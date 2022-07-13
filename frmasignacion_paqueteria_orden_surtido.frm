VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmasignacion_paqueteria_orden_surtido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de paqueteria y guias a ordenes de surtido"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   6975
   End
   Begin VB.Frame Frame5 
      Caption         =   " Por rango de fechas "
      Height          =   750
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   4425
      Begin VB.CommandButton cmd_ejecutar_filtro 
         Height          =   390
         Left            =   3855
         Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ejecuta filtro"
         Top             =   225
         Width           =   435
      End
      Begin VB.TextBox txt_inicio 
         Height          =   375
         Left            =   690
         TabIndex        =   10
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox txt_fin 
         Height          =   375
         Left            =   2505
         TabIndex        =   9
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2175
         TabIndex        =   11
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6390
      Left            =   60
      TabIndex        =   0
      Top             =   885
      Width           =   11490
      Begin VB.Frame frm_asigna_camion 
         Height          =   1845
         Left            =   2145
         TabIndex        =   15
         Top             =   1755
         Width           =   7155
         Begin VB.Frame frm_lista 
            Height          =   1845
            Left            =   1035
            TabIndex        =   16
            Top             =   15
            Width           =   5685
            Begin MSComctlLib.ListView lv_lista 
               Height          =   1365
               Left            =   30
               TabIndex        =   17
               Top             =   420
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   2408
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
                  Text            =   "Clave"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Nombre"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Cubicaje"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label Label11 
               BackColor       =   &H8000000D&
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   30
               TabIndex        =   18
               Top             =   135
               Width           =   5610
            End
         End
         Begin VB.TextBox txt_paqueteria 
            Height          =   360
            Left            =   1140
            TabIndex        =   20
            Top             =   795
            Width           =   1425
         End
         Begin VB.TextBox txt_nombre_paqueteria 
            Height          =   360
            Left            =   2595
            TabIndex        =   22
            Top             =   795
            Width           =   4410
         End
         Begin VB.TextBox txt_guia 
            Height          =   360
            Left            =   1140
            TabIndex        =   24
            Top             =   1185
            Width           =   2775
         End
         Begin VB.Frame Frame6 
            Height          =   60
            Left            =   15
            TabIndex        =   19
            Top             =   615
            Width           =   7095
         End
         Begin VB.CommandButton cmd_aceptar_pedidos 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   45
            Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":0192
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Aceptar Alt + A"
            Top             =   300
            Width           =   330
         End
         Begin VB.CommandButton cmd_cancelar_pedidos 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   375
            Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":02DC
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Cancelar Alt + C"
            Top             =   300
            Width           =   330
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   15
            Width           =   7140
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Paquetería:"
            Height          =   195
            Left            =   210
            TabIndex        =   23
            Top             =   885
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Guia:"
            Height          =   195
            Left            =   195
            TabIndex        =   21
            Top             =   1275
            Width           =   375
         End
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":0426
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":063C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar (Enter)"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   105
         Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":0958
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmasignacion_paqueteria_orden_surtido.frx":0A5A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   435
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_ordenes_surtido 
         Height          =   5520
         Left            =   105
         TabIndex        =   6
         Top             =   765
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   9737
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Orden"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre Agente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   2207
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre Cliente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Vol."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Paqueteria"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Guia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "paqueteria"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Ordenes de surtido "
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   7
         Top             =   135
         Width           =   11415
      End
   End
End
Attribute VB_Name = "frmasignacion_paqueteria_orden_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
  If Me.txt_nombre_paqueteria <> "" Then
     If Me.txt_guia <> "" Then
        var_j = 1
        For var_j = 1 To lv_ordenes_surtido.ListItems.Count
            Me.lv_ordenes_surtido.ListItems.Item(var_j).Selected = True
            If Me.lv_ordenes_surtido.selectedItem.SubItems(8) = "*" Then
               rs.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "', VCHA_PAQ_NOMBRE = '" + Me.txt_nombre_paqueteria + "', VCHA_PAQ_GUIA = '" + Me.txt_guia + "' WHERE INTE_ORS_ORDEN_SURTIDO = " + Me.lv_ordenes_surtido.selectedItem, cnn, adOpenDynamic, adLockOptimistic
               rs.Open "update tb_detalle_Cajas set  VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "', VCHA_PAQ_GUIA = '" + Me.txt_guia + "' where inte_ors_orden_surtido = " + Me.lv_ordenes_surtido.selectedItem, cnn, adOpenDynamic, adLockOptimistic
               Me.lv_ordenes_surtido.selectedItem.SubItems(6) = Me.txt_nombre_paqueteria
               Me.lv_ordenes_surtido.selectedItem.SubItems(7) = Me.txt_guia
               Me.lv_ordenes_surtido.selectedItem.SubItems(9) = Me.txt_paqueteria
            End If
        Next var_j
        For var_j = 1 To lv_ordenes_surtido.ListItems.Count
            lv_ordenes_surtido.ListItems(var_j).Selected = True
            lv_ordenes_surtido.selectedItem.SubItems(8) = ""
            lv_ordenes_surtido.ListItems.Item(var_j).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(1).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(2).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(3).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(4).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(5).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(6).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(7).Bold = False
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
            lv_ordenes_surtido.ListItems.Item(var_j).ListSubItems(7).ForeColor = &H80000012
        
        
        
        Next var_j
        
        
        Me.frm_asigna_camion.Visible = False
        MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
     Else
        MsgBox "Debe de indicar una guia", vbOKOnly, "ATENCION"
     End If
  Else
     MsgBox "Denbe de indicr una paqueteria", vbOKOnlyl, "ATENCION"
  End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Me.frm_asigna_camion.Visible = False
End Sub

Private Sub cmd_ejecutar_filtro_Click()
   Dim list_item As ListItem
   Me.lv_ordenes_surtido.ListItems.Clear
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         cnn.CommandTimeout = 6000
         var_dia = CStr(Day(CDate(Me.txt_inicio)))
         var_mes = CStr(Month(CDate(txt_inicio)))
         var_año = CStr(Year(CDate(txt_inicio)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         
         
         var_dia = CStr(Day(CDate(Me.txt_fin) + 1))
         var_mes = CStr(Month(CDate(Me.txt_fin) + 1))
         var_año = CStr(Year(CDate(Me.txt_fin)) + 1)
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         var_cadena = "SELECT  TOP 100 PERCENT SUM(dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.FLOA_LIN_VOLUMEN) AS VOLUMEN, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') AS VCHA_ORS_ESTATUS_CAMION, dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE , dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, TB_ENC_ORDEN_SURTIDO.vcha_paq_nombre, vcha_paq_guia, vcha_paq_clave_id  FROM dbo.VW_CLIENTES INNER JOIN  dbo.TB_AGENTES ON dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO ON"
         var_cadena = var_cadena + " dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + ") AND (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin + " - .000001) GROUP BY dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA'), dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID, TB_ENC_ORDEN_SURTIDO.vcha_paq_nombre, vcha_paq_guia, vcha_paq_clave_id ORDER BY dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
               list_item.SubItems(2) = IIf(IsNull(rs!VCHA_age_NOMBRE), "", rs!VCHA_age_NOMBRE)
               list_item.SubItems(3) = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
               list_item.SubItems(4) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
               list_item.SubItems(5) = Format(IIf(IsNull(rs!volumen), 0, rs!volumen), "###,###,##0.00")
               list_item.SubItems(6) = IIf(IsNull(rs!VCHA_PAQ_NOMBRE), "", rs!VCHA_PAQ_NOMBRE)
               list_item.SubItems(7) = IIf(IsNull(rs!vcha_paq_guia), "", rs!vcha_paq_guia)
               list_item.SubItems(9) = IIf(IsNull(rs!VCHA_PAQ_CLAVE_ID), "", rs!VCHA_PAQ_CLAVE_ID)
               rs.MoveNext:
         Wend
         rs.Close
         If Me.lv_ordenes_surtido.ListItems.Count > 17 Then
            Me.lv_ordenes_surtido.ColumnHeaders(5).Width = 2480.28
         Else
            Me.lv_ordenes_surtido.ColumnHeaders(5).Width = 2700.28
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_ordenes_surtido.ListItems.Count
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      If lv_ordenes_surtido.selectedItem.SubItems(8) = "*" Then
         lv_ordenes_surtido.selectedItem.SubItems(8) = ""
         lv_ordenes_surtido.ListItems.Item(i).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
      Else
         lv_ordenes_surtido.selectedItem.SubItems(8) = "*"
         lv_ordenes_surtido.ListItems.Item(i).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_ordenes_surtido.selectedItem.Index
   If lv_ordenes_surtido.selectedItem.SubItems(8) = "*" Then
      lv_ordenes_surtido.selectedItem.SubItems(8) = ""
      lv_ordenes_surtido.ListItems.Item(i).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
      lv_ordenes_surtido.Refresh
   Else
      lv_ordenes_surtido.selectedItem.SubItems(8) = "*"
      lv_ordenes_surtido.ListItems.Item(i).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      lv_ordenes_surtido.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_ordenes_surtido.ListItems.Count
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      lv_ordenes_surtido.selectedItem.SubItems(8) = ""
      lv_ordenes_surtido.ListItems.Item(i).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
   Next i
   lv_ordenes_surtido.Refresh
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_ordenes_surtido.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_ordenes_surtido.selectedItem.SubItems(8) = "" And var_rellena = True Then
            lv_ordenes_surtido.selectedItem.SubItems(8) = "*"
            lv_ordenes_surtido.ListItems.Item(i).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_ordenes_surtido.selectedItem.SubItems(8) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_ordenes_surtido.selectedItem.SubItems(8) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_ordenes_surtido.ListItems.Count
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      lv_ordenes_surtido.selectedItem.SubItems(8) = "*"
      lv_ordenes_surtido.ListItems.Item(i).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = True
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
   Next i
   lv_ordenes_surtido.Refresh
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Me.txt_fin = Date
   Me.txt_inicio = Date
   Me.frm_asigna_camion.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_paqueteria = lv_lista.selectedItem
      Me.txt_nombre_paqueteria = lv_lista.selectedItem.SubItems(1)
      Me.txt_paqueteria.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_ordenes_surtido_Click()
   Me.frm_asigna_camion.Visible = False
End Sub

Private Sub lv_ordenes_surtido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Me.frm_asigna_camion.Visible = False
   Call pro_ordena_listas(Me.lv_ordenes_surtido, ColumnHeader)
End Sub

Private Sub lv_ordenes_surtido_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   Me.frm_asigna_camion.Visible = False
End Sub

Private Sub lv_ordenes_surtido_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.frm_asigna_camion.Visible = False
End Sub

Private Sub lv_ordenes_surtido_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.frm_asigna_camion.Visible = False
   If KeyCode = 116 Then
      Me.frm_asigna_camion.Visible = True
      Me.frm_lista.Visible = False
      Me.txt_nombre_paqueteria = ""
      Me.txt_paqueteria = ""
      Me.txt_guia = ""
      Me.txt_paqueteria.SetFocus
   End If
End Sub

Private Sub lv_ordenes_surtido_KeyPress(KeyAscii As Integer)
   Me.frm_asigna_camion.Visible = False
   If KeyAscii = 13 Then
      i = lv_ordenes_surtido.selectedItem.Index
      If lv_ordenes_surtido.selectedItem.SubItems(8) = "*" Then
         lv_ordenes_surtido.selectedItem.SubItems(8) = ""
         lv_ordenes_surtido.ListItems.Item(i).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
         lv_ordenes_surtido.Refresh
      Else
         lv_ordenes_surtido.selectedItem.SubItems(8) = "*"
         lv_ordenes_surtido.ListItems.Item(i).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_ordenes_surtido.Refresh
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      Me.txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_guia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      Me.txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_transporte_Change()

End Sub

Private Sub txt_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_PAQUETERIA order by vcha_paq_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAQ_CLAVE_ID, "", rs!VCHA_PAQ_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PAQ_NOMBRE), "", rs!VCHA_PAQ_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAQUETERIAS"
      var_tipo_lista = 101
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

Private Sub txt_nombre_paqueteria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_guia.SetFocus
   Else
      If KeyAscii = 27 Then
         Me.frm_asigna_camion.Visible = False
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_paqueteria_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_lista = 1
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_PAQUETERIA order by vcha_PAQ_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAQ_CLAVE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PAQ_NOMBRE), "", rs!VCHA_PAQ_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      Label11 = "PAQUETERIA"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      lv_lista.ColumnHeaders(2).Width = 3000
      lv_lista.ColumnHeaders(3).Width = 1350
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_paqueteria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_paqueteria.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
End Sub

Private Sub txt_paqueteria_LostFocus()
   If Trim(Me.txt_paqueteria) <> "" Then
      rs.Open "select * from tb_paqueteria where vcha_paq_clave_id = '" + Me.txt_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_paqueteria = IIf(IsNull(rs!VCHA_PAQ_NOMBRE), "", rs!VCHA_PAQ_NOMBRE)
      Else
         Me.txt_nombre_paqueteria = ""
         Me.txt_paqueteria = ""
         MsgBox "Paqueteria incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_paqueteria = ""
   End If
End Sub
