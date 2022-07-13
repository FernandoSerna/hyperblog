VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmordenes_surtido_asignacion_camion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de camion, chofer, trajunante y volumen a las ordenes de surtido"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_asigna_camion 
      Height          =   2565
      Left            =   2505
      TabIndex        =   22
      Top             =   3105
      Width           =   7155
      Begin VB.Frame frm_lista_2 
         Height          =   2400
         Left            =   1005
         TabIndex        =   39
         Top             =   60
         Width           =   5685
         Begin MSComctlLib.ListView lv_lista_2 
            Height          =   1905
            Left            =   30
            TabIndex        =   40
            Top             =   450
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   3360
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
            TabIndex        =   41
            Top             =   135
            Width           =   5610
         End
      End
      Begin VB.CommandButton cmd_cancelar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   300
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   300
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   60
         Left            =   15
         TabIndex        =   36
         Top             =   615
         Width           =   7095
      End
      Begin VB.TextBox txt_volumen_asignado 
         Height          =   360
         Left            =   5235
         TabIndex        =   28
         Top             =   1185
         Width           =   1770
      End
      Begin VB.TextBox txt_trajinante 
         Height          =   360
         Left            =   1140
         TabIndex        =   31
         Top             =   1980
         Width           =   5865
      End
      Begin VB.TextBox txt_nombre_chofer 
         Height          =   360
         Left            =   2595
         TabIndex        =   30
         Top             =   1575
         Width           =   4410
      End
      Begin VB.TextBox txt_chofer 
         Height          =   360
         Left            =   1140
         TabIndex        =   29
         Top             =   1575
         Width           =   1425
      End
      Begin VB.TextBox txt_volumen_transporte 
         Height          =   360
         Left            =   1905
         TabIndex        =   27
         Top             =   1185
         Width           =   1770
      End
      Begin VB.TextBox txt_nombre_transporte 
         Height          =   360
         Left            =   2595
         TabIndex        =   26
         Top             =   795
         Width           =   4410
      End
      Begin VB.TextBox txt_transporte 
         Height          =   360
         Left            =   1140
         TabIndex        =   25
         Top             =   795
         Width           =   1425
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Volumen asignado:"
         Height          =   195
         Left            =   3825
         TabIndex        =   35
         Top             =   1275
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Trajinante:"
         Height          =   195
         Left            =   225
         TabIndex        =   34
         Top             =   2070
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Chofer:"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   1665
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Volumen del transporte:"
         Height          =   195
         Left            =   195
         TabIndex        =   32
         Top             =   1275
         Width           =   1665
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   885
         Width           =   810
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   15
         Width           =   7140
      End
   End
   Begin VB.Frame frm_eliminar_asignacion 
      Height          =   2955
      Left            =   2490
      TabIndex        =   42
      Top             =   2820
      Width           =   7035
      Begin MSComctlLib.ListView lv_eliminar_asignacion 
         Height          =   2460
         Left            =   45
         TabIndex        =   44
         Top             =   405
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4339
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Transporte"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Chofer"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "V. Transporte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "V. Asignado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Consecutivo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_eliminar_asignacion 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   45
         TabIndex        =   43
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   3945
      TabIndex        =   19
      Top             =   150
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   20
         Top             =   495
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
         TabIndex        =   21
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1365
      Left            =   60
      TabIndex        =   12
      Top             =   15
      Width           =   11490
      Begin VB.CommandButton cmd_ejecutar_filtro 
         Height          =   390
         Left            =   10800
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ejecuta filtro"
         Top             =   690
         Width           =   435
      End
      Begin VB.Frame Frame5 
         Caption         =   " Por rango de fechas "
         Height          =   750
         Left            =   6015
         TabIndex        =   15
         Top             =   510
         Width           =   4695
         Begin VB.TextBox txt_fin 
            Height          =   375
            Left            =   2955
            TabIndex        =   3
            Top             =   225
            Width           =   1230
         End
         Begin VB.TextBox txt_inicio 
            Height          =   375
            Left            =   1140
            TabIndex        =   2
            Top             =   225
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fin:"
            Height          =   195
            Left            =   2625
            TabIndex        =   17
            Top             =   315
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   675
            TabIndex        =   16
            Top             =   315
            Width           =   420
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Por agente "
         Height          =   750
         Left            =   105
         TabIndex        =   14
         Top             =   510
         Width           =   5835
         Begin VB.TextBox txt_nombre_agente 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   270
            Width           =   4875
         End
         Begin VB.TextBox txt_clave_agente 
            Height          =   375
            Left            =   90
            TabIndex        =   0
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Filtrado de ordenes de surtido"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   11415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5850
      Left            =   60
      TabIndex        =   11
      Top             =   1365
      Width           =   11490
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":0426
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   105
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":063C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":073E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar (Enter)"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmordenes_surtido_asignacion_camion.frx":0A5A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   435
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_ordenes_surtido 
         Height          =   5010
         Left            =   105
         TabIndex        =   5
         Top             =   765
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   8837
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre Agente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   2207
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre Cliente"
            Object.Width           =   4763
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Volumen"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estatus"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Ruta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Nombre Ruta"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Ordenes de surtido "
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   18
         Top             =   135
         Width           =   11415
      End
   End
End
Attribute VB_Name = "frmordenes_surtido_asignacion_camion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_lista As Integer
Private Sub cmd_aceptar_pedidos_Click()
   Dim var_posible  As Boolean
   var_posible = False
   If Me.lv_ordenes_surtido.ListItems.Count > 0 Then
      For var_j = 1 To lv_ordenes_surtido.ListItems.Count
          Me.lv_ordenes_surtido.ListItems.Item(var_j).Selected = True
          If Me.lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
             var_posible = True
          End If
      Next var_j
   End If
   If var_posible = True Then
      If Me.txt_transporte <> "" Then
         If Me.txt_chofer <> "" Then
            If IsNumeric(Me.txt_volumen_transporte) Then
               If IsNumeric(Me.txt_volumen_asignado) Then
                  If Me.txt_trajinante <> "" Then
                     var_j = 1
                     For var_j = 1 To lv_ordenes_surtido.ListItems.Count
                         Me.lv_ordenes_surtido.ListItems.Item(var_j).Selected = True
                         If Me.lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
                         
                            var_cadena = "INSERT INTO TB_ASIGNACION_ORDEN_SURTIDO_CAMION (INTE_ORS_ORDEN_SURTIDO, FLOA_ORS_VOLUMEN_ORDEN, FLOA_ORS_VOLUMEN_ORDEN_CAMION, FLOA_ORS_VOLUMEN_ASIGNADO, VCHA_TRA_TRANSPORTE_ID, VCHA_TRA_NOMBRE, VCHA_CHO_CHOFER_ID, VCHA_CHO_NOMBRE, VCHA_ORS_TRAJINANTE, VCHA_RUT_RUTA_ID) Values"
                            var_cadena = var_cadena + "(" + Me.lv_ordenes_surtido.selectedItem + ", " + Me.lv_ordenes_surtido.selectedItem.SubItems(5) + "," + Me.txt_volumen_transporte + ", " + Me.txt_volumen_asignado + ",'" + Me.txt_transporte + "','" + Me.txt_nombre_transporte + "', '" + Me.txt_chofer + "','" + Me.txt_nombre_chofer + "', '" + Me.txt_trajinante + "', '" + Me.lv_ordenes_surtido.selectedItem.SubItems(8) + "')"
                            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                            
                            rs.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET VCHA_ORS_ESTATUS_CAMION = 'ASIGNADA', FLOA_ORS_VOLUMEN_ORDEN = " + Me.lv_ordenes_surtido.selectedItem.SubItems(5) + ", FLOA_ORS_VOLUMEN_ORDEN_CAMION = " + Me.txt_volumen_transporte + ", VCHA_TRA_TRANSPORTE_ID = '" + Me.txt_transporte + "', VCHA_TRA_NOMBRE = '" + Me.txt_nombre_transporte + "', VCHA_CHO_CHOFER_ID = '" + Me.txt_chofer + "' , VCHA_CHO_NOMBRE = '" + Me.txt_nombre_chofer + "', VCHA_ORS_TRAJINANTE = '" + Me.txt_trajinante + "', VCHA_RUT_RUTA_ID = '" + Me.lv_ordenes_surtido.selectedItem.SubItems(8) + "', INTE_ORS_AÑO = " + CStr(Year(Date)) + ", INTE_ORS_MES = " + CStr(Month(Date)) + ",INTE_ORS_DIA = " + CStr(Day(Date)) + " WHERE INTE_ORS_ORDEN_SURTIDO = " + Me.lv_ordenes_surtido.selectedItem, cnn, adOpenDynamic, adLockOptimistic
                            Me.lv_ordenes_surtido.selectedItem.SubItems(6) = "ASIGNADA"
                         End If
                     Next var_j
                     MsgBox "Se a terminado el proceso de asignación de ordenes de surtido", vbOKOnly, "ATENCION"
                     
                     
                     n = lv_ordenes_surtido.ListItems.Count
                     For i = 1 To n
                         lv_ordenes_surtido.ListItems.Item(i).Selected = True
                         lv_ordenes_surtido.selectedItem.SubItems(7) = ""
                         lv_ordenes_surtido.ListItems.Item(i).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
                         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
                     Next i
                     lv_ordenes_surtido.Refresh
                     
                     
                     
                     Me.frm_asigna_camion.Visible = False
                     Me.lv_ordenes_surtido.SetFocus
                  Else
                     MsgBox "Debe de indicar un trajinante", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Volumen asignado incorrecto", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Volumen de transporte incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se debe de seleccionar un chofer", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Se debe de seleccionar un transporte", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
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
         
         
         
         If Trim(Me.txt_clave_agente) <> "" Then
            var_cadena = "SELECT  TOP 100 PERCENT SUM(dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.FLOA_LIN_VOLUMEN) AS VOLUMEN, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') AS VCHA_ORS_ESTATUS_CAMION, dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE , dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID FROM dbo.VW_CLIENTES INNER JOIN  dbo.TB_AGENTES ON dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO ON"
            var_cadena = var_cadena + " dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + ") AND (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin + " - .000001) AND (dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID = '" + Me.txt_clave_agente + "') GROUP BY dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA'), dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID ORDER BY dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO "
            'var_cadena = "SELECT TOP 100 PERCENT SUM(dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.FLOA_LIN_VOLUMEN) AS VOLUMEN, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') AS VCHA_ORS_ESTATUS_CAMION FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN ON"
            'var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + ") AND (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin + " -.000001) and dbo.tb_agentes.vcha_age_agente_id = '" + Me.txt_clave_agente + "'  GROUP BY dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') ORDER BY dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO   "
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
                  list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_CLAVE_ID), "", rs!VCHA_CLI_CLAVE_ID)
                  list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!volumen), 0, rs!volumen), "###,###,##0.00")
                  list_item.SubItems(6) = IIf(IsNull(rs!VCHA_ORS_ESTATUS_CAMION), "NO ASIGNADA", rs!VCHA_ORS_ESTATUS_CAMION)
                  list_item.SubItems(8) = IIf(IsNull(rs!vcha_rut_ruta_id), "", rs!vcha_rut_ruta_id)
                  list_item.SubItems(9) = IIf(IsNull(rs!VCHA_RUT_NOMBRE), "", rs!VCHA_RUT_NOMBRE)
                  rs.MoveNext:
            Wend
            rs.Close
         Else
            var_cadena = "SELECT  TOP 100 PERCENT SUM(dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.FLOA_LIN_VOLUMEN) AS VOLUMEN, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') AS VCHA_ORS_ESTATUS_CAMION, dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE , dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID FROM dbo.VW_CLIENTES INNER JOIN  dbo.TB_AGENTES ON dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO ON"
            var_cadena = var_cadena + " dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + ") AND (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin + " - .000001) GROUP BY dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA'), dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.VW_CLIENTES.VCHA_CLI_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_NOMBRE, dbo.VW_CLIENTES.VCHA_RUT_RUTA_ID ORDER BY dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO "
            'var_cadena = "SELECT TOP 100 PERCENT SUM(dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.FLOA_LIN_VOLUMEN) AS VOLUMEN, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') AS VCHA_ORS_ESTATUS_CAMION FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN ON"
            'var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + ") AND (dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin + " -.000001) GROUP BY dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, ISNULL(dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ORS_ESTATUS_CAMION, 'NO ASIGNADA') ORDER BY dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_DET_ORDEN_SURTIDO_VOLUMEN.INTE_ORS_ORDEN_SURTIDO   "
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
                  list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_CLAVE_ID), "", rs!VCHA_CLI_CLAVE_ID)
                  list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!volumen), 0, rs!volumen), "###,###,##0.00")
                  list_item.SubItems(6) = IIf(IsNull(rs!VCHA_ORS_ESTATUS_CAMION), "NO ASIGNADA", rs!VCHA_ORS_ESTATUS_CAMION)
                  list_item.SubItems(8) = IIf(IsNull(rs!vcha_rut_ruta_id), "", rs!vcha_rut_ruta_id)
                  list_item.SubItems(9) = IIf(IsNull(rs!VCHA_RUT_NOMBRE), "", rs!VCHA_RUT_NOMBRE)
                  rs.MoveNext:
            Wend
            rs.Close
         End If
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
      If lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
         lv_ordenes_surtido.selectedItem.SubItems(7) = ""
         lv_ordenes_surtido.ListItems.Item(i).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      Else
         If Me.lv_ordenes_surtido.selectedItem.SubItems(6) <> "ASIGNADA" Then
            lv_ordenes_surtido.selectedItem.SubItems(7) = "*"
            lv_ordenes_surtido.ListItems.Item(i).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         End If
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_ordenes_surtido.selectedItem.Index
   If lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
      lv_ordenes_surtido.selectedItem.SubItems(7) = ""
      lv_ordenes_surtido.ListItems.Item(i).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_ordenes_surtido.Refresh
   Else
      If Me.lv_ordenes_surtido.selectedItem.SubItems(6) <> "ASIGNADA" Then
         lv_ordenes_surtido.selectedItem.SubItems(7) = "*"
         lv_ordenes_surtido.ListItems.Item(i).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_ordenes_surtido.Refresh
      End If
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_ordenes_surtido.ListItems.Count
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      lv_ordenes_surtido.selectedItem.SubItems(7) = ""
      lv_ordenes_surtido.ListItems.Item(i).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
   Next i
   lv_ordenes_surtido.Refresh
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_ordenes_surtido.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_ordenes_surtido.selectedItem.SubItems(7) = "" And var_rellena = True Then
         If Me.lv_ordenes_surtido.selectedItem.SubItems(6) <> "ASIGNADA" Then
            lv_ordenes_surtido.selectedItem.SubItems(7) = "*"
            lv_ordenes_surtido.ListItems.Item(i).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         End If
      Else
         If var_encontro = True And lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_ordenes_surtido.selectedItem.SubItems(7) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_ordenes_surtido.ListItems.Count
   For i = 1 To n
      lv_ordenes_surtido.ListItems.Item(i).Selected = True
      If Me.lv_ordenes_surtido.selectedItem.SubItems(6) <> "ASIGNADA" Then
         lv_ordenes_surtido.selectedItem.SubItems(7) = "*"
         lv_ordenes_surtido.ListItems.Item(i).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      End If
   Next i
   lv_ordenes_surtido.Refresh
End Sub

Private Sub Form_Load()
   Me.txt_inicio = Date
   Me.txt_fin = Date
   Me.frm_lista.Visible = False
   Me.frm_asigna_camion.Visible = False
   Me.frm_lista_2.Visible = False
   Me.frm_eliminar_asignacion.Visible = False
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lv_eliminar_asignacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_si = MsgBox("¿Desea eliminar la asignacion?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación de la asignación", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "DELETE FROM TB_ASIGNACION_ORDEN_SURTIDO_CAMION WHERE INTE_ORS_CONSECUTIVO = " + Me.lv_eliminar_asignacion.selectedItem.SubItems(4)
            Me.lv_eliminar_asignacion.ListItems.Clear
            rs.Open "select * from TB_ASIGNACION_ORDEN_SURTIDO_CAMION where inte_ors_orden_surtido = " + CStr(CDbl(Me.lv_ordenes_surtido.selectedItem)), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = Me.lv_eliminar_asignacion.ListItems.Add(, , rs!vcha_tra_nombre)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_cho_nombre), "", rs!vcha_cho_nombre)
                     list_item.SubItems(2) = IIf(IsNull(rs!FLOA_ORS_VOLUMEN_ORDEN_CAMION), "", rs!FLOA_ORS_VOLUMEN_ORDEN_CAMION)
                     list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ORS_VOLUMEN_ASIGNADO), "", rs!FLOA_ORS_VOLUMEN_ASIGNADO)
                     list_item.SubItems(4) = IIf(IsNull(rs!INTE_ORS_CONSECUTIVO), "", rs!INTE_ORS_CONSECUTIVO)
                     rs.MoveNext
               Wend
               rs.Close
            Else
               rs.Close
               rs.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET VCHA_ORS_ESTATUS_CAMION = 'NO ASIGNADA' WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(CDbl(Me.lv_ordenes_surtido.selectedItem)), cnn, adOpenDynamic, adLockOptimistic
               Me.lv_ordenes_surtido.selectedItem.SubItems(6) = "NO ASIGNADA"
            End If
         End If
      End If
      Me.lv_eliminar_asignacion.ListItems.Clear
      Me.frm_eliminar_asignacion.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_eliminar_asignacion.Visible = False
   End If
End Sub

Private Sub lv_eliminar_asignacion_LostFocus()
   Me.lv_eliminar_asignacion.ListItems.Clear
   Me.frm_eliminar_asignacion.Visible = False
End Sub

Private Sub lv_lista_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_lista = 1 Then
         Me.txt_transporte = Me.lv_lista_2.selectedItem
         Me.txt_nombre_transporte = Me.lv_lista_2.selectedItem.SubItems(1)
         Me.txt_transporte.SetFocus
      End If
      If var_lista = 2 Then
         Me.txt_chofer = Me.lv_lista_2.selectedItem
         Me.txt_nombre_chofer = Me.lv_lista_2.selectedItem.SubItems(1)
         Me.txt_chofer.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_lista = 1 Then
         Me.txt_transporte.SetFocus
      End If
      If var_lista = 2 Then
         Me.txt_chofer.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_2_LostFocus()
    Me.frm_lista_2.Visible = False
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_clave_agente = lv_lista.selectedItem
      Me.txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
      Me.txt_clave_agente.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_clave_agente.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_ordenes_surtido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_ordenes_surtido, ColumnHeader)
End Sub

Private Sub lv_ordenes_surtido_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_cubicaje = 0
      For var_i = 1 To Me.lv_ordenes_surtido.ListItems.Count
          Me.lv_ordenes_surtido.ListItems(var_i).Selected = True
          If Me.lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
             var_cubicaje = var_cubicaje + CDbl(Me.lv_ordenes_surtido.selectedItem.SubItems(5))
          End If
      Next var_i
      Me.txt_chofer = ""
      Me.txt_nombre_chofer = ""
      Me.txt_nombre_transporte = ""
      Me.txt_trajinante = ""
      Me.txt_transporte = ""
      Me.txt_volumen_asignado = Format(var_cubicaje, "###,###,##0.00")
      Me.txt_volumen_transporte = ""
      Me.frm_asigna_camion.Visible = True
      Me.txt_transporte.SetFocus
   End If
   If KeyCode = 114 Then
      Me.lv_eliminar_asignacion.ListItems.Clear
      rs.Open "select * from TB_ASIGNACION_ORDEN_SURTIDO_CAMION where inte_ors_orden_surtido = " + CStr(CDbl(Me.lv_ordenes_surtido.selectedItem)), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = Me.lv_eliminar_asignacion.ListItems.Add(, , rs!vcha_tra_nombre)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_cho_nombre), "", rs!vcha_cho_nombre)
               list_item.SubItems(2) = IIf(IsNull(rs!FLOA_ORS_VOLUMEN_ORDEN_CAMION), "", rs!FLOA_ORS_VOLUMEN_ORDEN_CAMION)
               list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ORS_VOLUMEN_ASIGNADO), "", rs!FLOA_ORS_VOLUMEN_ASIGNADO)
               list_item.SubItems(4) = IIf(IsNull(rs!INTE_ORS_CONSECUTIVO), "", rs!INTE_ORS_CONSECUTIVO)
               rs.MoveNext
         Wend
         rs.Close
      
         Me.lbl_eliminar_asignacion = "Orden de Surtido " + Me.lv_ordenes_surtido.selectedItem
         var_tipo_lista = 1
         Me.frm_eliminar_asignacion.Visible = True
         Me.lv_eliminar_asignacion.SetFocus
      Else
         MsgBox "No se a asignado la orden", vbOKOnly, "ATENCION"
         rs.Close
      End If
   End If
   
End Sub

Private Sub lv_ordenes_surtido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_ordenes_surtido.selectedItem.Index
      If lv_ordenes_surtido.selectedItem.SubItems(7) = "*" Then
         lv_ordenes_surtido.selectedItem.SubItems(7) = ""
         lv_ordenes_surtido.ListItems.Item(i).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_ordenes_surtido.Refresh
      Else
         If Me.lv_ordenes_surtido.selectedItem.SubItems(6) <> "ASIGNADA" Then
            lv_ordenes_surtido.selectedItem.SubItems(7) = "*"
            lv_ordenes_surtido.ListItems.Item(i).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_ordenes_surtido.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_ordenes_surtido.Refresh
         End If
      End If
   End If
End Sub

Private Sub txt_chofer_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_lista = 2
      lv_lista_2.ListItems.Clear
      rs.Open "select * from tb_choferes order by vcha_cho_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista_2.ListItems.Add(, , rs!vcha_cho_chofer_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cho_nombre), "", rs!vcha_cho_nombre)
            rs.MoveNext
      Wend
      rs.Close
      Label11 = "CHOFERES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista_2.ListItems.Count
      lv_lista_2.ColumnHeaders(2).Width = 4300
      lv_lista_2.ColumnHeaders(3).Width = 0
      frm_lista_2.Visible = True
      lv_lista_2.SetFocus
   End If
End Sub

Private Sub txt_chofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_chofer_LostFocus()
    If Trim(Me.txt_chofer) <> "" Then
       rs.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + Me.txt_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          Me.txt_nombre_chofer = IIf(IsNull(rs!vcha_cho_nombre), "", rs!vcha_cho_nombre)
       Else
          Me.txt_chofer = ""
          Me.txt_nombre_chofer = ""
          MsgBox "Clave de chofer incorrecta", vbOKOnly, "ATENCION"
       End If
       rs.Close
    Else
       Me.txt_nombre_chofer = ""
    End If
End Sub

Private Sub txt_clave_agente_Change()
   Me.txt_nombre_agente = ""
End Sub

Private Sub txt_clave_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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

Private Sub txt_clave_agente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_agente_LostFocus()
   If Me.txt_clave_agente = "" Then
      Me.txt_nombre_agente = ""
   Else
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + Me.txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_clave_agente = ""
         Me.txt_nombre_agente = ""
      End If
      rs.Close
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

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fin_LostFocus()
   If Not IsDate(Me.txt_fin) Then
      MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      Me.txt_fin = Date
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

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   If Not IsDate(Me.txt_inicio) Then
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
      Me.txt_inicio = Date
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_chofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_transporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_trajinante_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_lista = 1
      lv_lista_2.ListItems.Clear
      rs.Open "select * from tb_transportes order by vcha_trN_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista_2.ListItems.Add(, , rs!VCHA_TRN_TRANSPORTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_trn_cubicaje), 0, rs!floa_trn_cubicaje), "###,###,##0.00")
            rs.MoveNext
      Wend
      rs.Close
      Label11 = "TRANSPORTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista_2.ListItems.Count
      lv_lista_2.ColumnHeaders(2).Width = 3000
      lv_lista_2.ColumnHeaders(3).Width = 1350
      frm_lista_2.Visible = True
      lv_lista_2.SetFocus
   End If

End Sub

Private Sub txt_transporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_transporte_LostFocus()
   If Trim(Me.txt_transporte) <> "" Then
       rs.Open "select * from tb_transportes where vcha_trn_transporte_id = '" + Me.txt_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          Me.txt_nombre_transporte = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
          Me.txt_volumen_transporte = IIf(IsNull(rs!floa_trn_cubicaje), 0, rs!floa_trn_cubicaje)
       Else
          Me.txt_nombre_transporte = ""
          Me.txt_transporte = ""
          MsgBox "Clave de transporte incorrecta", vbOKOnly, "ATENCION"
       End If
       rs.Close
   Else
      Me.txt_nombre_transporte = ""
   End If
End Sub

Private Sub txt_volumen_asignado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_volumen_transporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_asigna_camion.Visible = False
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub
