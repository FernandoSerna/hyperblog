VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_activar_rutas_distribucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activar dias festivos"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   -120
      Width           =   18855
      Begin VB.TextBox txt_establecimiento 
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   210
         Width           =   2895
      End
      Begin VB.ComboBox cmb_dias 
         Height          =   315
         ItemData        =   "frmoracle_activar_rutas_distribucion.frx":0000
         Left            =   120
         List            =   "frmoracle_activar_rutas_distribucion.frx":0019
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         Picture         =   "frmoracle_activar_rutas_distribucion.frx":0059
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar establecimiento:"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8145
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   18915
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_activar_rutas_distribucion.frx":015B
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_activar_rutas_distribucion.frx":0371
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_activar_rutas_distribucion.frx":0473
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_activar_rutas_distribucion.frx":0545
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmoracle_activar_rutas_distribucion.frx":078F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   7575
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   18690
         _ExtentX        =   32967
         _ExtentY        =   13361
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
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ruta"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Titular"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Establecimiento"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lunes"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Martes"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Miercoles"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Jueves"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Viernes"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Sabado"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Domingo"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "L_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "M_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "M_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "J_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "V_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "S_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "D_Fest"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Establecimiento"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_activar_rutas_distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Me.cmb_dias <> "LUNES" And Me.cmb_dias <> "MARTES" And Me.cmb_dias <> "MIERCOLES" And Me.cmb_dias <> "JUEVES" And Me.cmb_dias <> "VIERNES" And Me.cmb_dias <> "SABADO" And Me.cmb_dias <> "DOMINGO" Then
      MsgBox "Dia incorrecto", vbOKOnly, "ATENCION"
   Else
      var_si = MsgBox("¿Desea desactivar el " + Me.cmb_dias + " como dia pedido?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la desactivacion del dia " + Me.cmb_dias, vbYesNo, "ATENCION")
         If var_si = 6 Then
            If Me.cmb_dias = "LUNES" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(4) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET LUNES_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            If Me.cmb_dias = "MARTES" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(5) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET MARTES_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            If Me.cmb_dias = "MIERCOLES" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(6) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET MIERCOLES_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            
            If Me.cmb_dias = "JUEVES" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(7) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET JUEVES_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            
            If Me.cmb_dias = "VIERNES" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(8) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET VIERNES_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            
            If Me.cmb_dias = "SABADO" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(9) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET SABADO_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            
            If Me.cmb_dias = "DOMINGO" Then
               For var_j = 1 To Me.lv_clientes.ListItems.Count
                   Me.lv_clientes.ListItems.Item(var_j).Selected = True
                   If Me.lv_clientes.selectedItem.SubItems(10) = "1" Then
                      If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                         rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET DOMINGO_FEST = 1 WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      End If
                   End If
               Next var_j
            End If
            
            
            
            Me.lv_clientes.ListItems.Clear
            
   'rs.Open "select * from xxvia_vw_dias_despacho  order by nombre_ruta, nombre_Establecimiento", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "select party_site_number, RUTA,NOMBRE_RUTA,a.site_use_id,TITULAR,NOMBRE_TITULAR, LUNES,MARTES,MIERCOLES,JUEVES,VIERNES,SABADO,DOMINGO,PAQUETERIA,LUNES_FEST,MARTES_FEST,MIERCOLES_FEST,JUEVES_FEST,VIERNES_FEST,SABADO_FEST,DOMINGO_FEST, party_site_number||'   '||a.NOMBRE_ESTABLECIMIENTO nombre_establecimiento from xxvia_vw_dias_despacho a, XXVIA_VW_CLIENTES_BCP b where a.site_use_id = to_char(b.site_use_id)  order by nombre_ruta, nombre_Establecimiento ", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_clientes.ListItems.Add(, , rs!nombre_ruta)
         list_item.SubItems(1) = IIf(IsNull(rs!nombre_titular), "", rs!nombre_titular)
         list_item.SubItems(2) = IIf(IsNull(rs!site_use_id), "", rs!site_use_id)
         list_item.SubItems(3) = IIf(IsNull(rs!nombre_Establecimiento), "", rs!nombre_Establecimiento)
         list_item.SubItems(4) = IIf(IsNull(rs!lunes), "", rs!lunes)
         list_item.SubItems(5) = IIf(IsNull(rs!martes), "", rs!martes)
         list_item.SubItems(6) = IIf(IsNull(rs!miercoles), "", rs!miercoles)
         list_item.SubItems(7) = IIf(IsNull(rs!jueves), "", rs!jueves)
         list_item.SubItems(8) = IIf(IsNull(rs!viernes), "", rs!viernes)
         list_item.SubItems(9) = IIf(IsNull(rs!sabado), "", rs!sabado)
         list_item.SubItems(10) = IIf(IsNull(rs!domingo), "", rs!domingo)
         list_item.SubItems(11) = IIf(IsNull(rs!lunes_fest), "", rs!lunes_fest)
         list_item.SubItems(12) = IIf(IsNull(rs!martes_fest), "", rs!martes_fest)
         list_item.SubItems(13) = IIf(IsNull(rs!miercoles_fest), "", rs!miercoles_fest)
         list_item.SubItems(14) = IIf(IsNull(rs!jueves_fest), "", rs!jueves_fest)
         list_item.SubItems(15) = IIf(IsNull(rs!viernes_fest), "", rs!viernes_fest)
         list_item.SubItems(16) = IIf(IsNull(rs!sabado_fest), "", rs!sabado_fest)
         list_item.SubItems(17) = IIf(IsNull(rs!domingo_fest), "", rs!domingo_fest)
         list_item.SubItems(18) = ""
         list_item.SubItems(19) = 0
         list_item.SubItems(20) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
         var_estatus = IIf(IsNull(rs!lunes_fest), 0, rs!lunes_fest) + IIf(IsNull(rs!martes_fest), 0, rs!martes_fest) + IIf(IsNull(rs!miercoles_fest), 0, rs!miercoles_fest) + IIf(IsNull(rs!jueves_fest), 0, rs!jueves_fest) + IIf(IsNull(rs!viernes_fest), 0, rs!viernes_fest) + IIf(IsNull(rs!sabado_fest), 0, rs!sabado_fest) + IIf(IsNull(rs!domingo_fest), 0, rs!domingo_fest)
         If var_estatus > 0 Then
            list_item.SubItems(19) = 1
         End If
         rs.MoveNext
   Wend
   rs.Close
   For var_i = 1 To Me.lv_clientes.ListItems.Count
       Me.lv_clientes.ListItems.Item(var_i).Selected = True
       If Me.lv_clientes.selectedItem.SubItems(19) = "1" Then
          lv_clientes.ListItems.Item(var_i).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).Bold = True
          lv_clientes.ListItems.Item(var_i).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).ForeColor = &HFF&
       Else
          lv_clientes.ListItems.Item(var_i).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).Bold = False
          lv_clientes.ListItems.Item(var_i).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).ForeColor = &H80000008
       End If
   Next var_i
            
            
            
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   Dim list_item As ListItem
   'rs.Open "select * from xxvia_vw_dias_despacho  order by nombre_ruta, nombre_Establecimiento", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "select party_site_number, RUTA,a.site_use_id,NOMBRE_RUTA,TITULAR,NOMBRE_TITULAR, LUNES,MARTES,MIERCOLES,JUEVES,VIERNES,SABADO,DOMINGO,PAQUETERIA,LUNES_FEST,MARTES_FEST,MIERCOLES_FEST,JUEVES_FEST,VIERNES_FEST,SABADO_FEST,DOMINGO_FEST, party_site_number||'   '||a.NOMBRE_ESTABLECIMIENTO nombre_establecimiento from xxvia_vw_dias_despacho a, XXVIA_VW_CLIENTES_BCP b where a.site_use_id = to_char(b.site_use_id)  order by nombre_ruta, nombre_Establecimiento ", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_clientes.ListItems.Add(, , rs!nombre_ruta)
         list_item.SubItems(1) = IIf(IsNull(rs!nombre_titular), "", rs!nombre_titular)
         list_item.SubItems(2) = IIf(IsNull(rs!site_use_id), "", rs!site_use_id)
         list_item.SubItems(3) = IIf(IsNull(rs!nombre_Establecimiento), "", rs!nombre_Establecimiento)
         list_item.SubItems(4) = IIf(IsNull(rs!lunes), "", rs!lunes)
         list_item.SubItems(5) = IIf(IsNull(rs!martes), "", rs!martes)
         list_item.SubItems(6) = IIf(IsNull(rs!miercoles), "", rs!miercoles)
         list_item.SubItems(7) = IIf(IsNull(rs!jueves), "", rs!jueves)
         list_item.SubItems(8) = IIf(IsNull(rs!viernes), "", rs!viernes)
         list_item.SubItems(9) = IIf(IsNull(rs!sabado), "", rs!sabado)
         list_item.SubItems(10) = IIf(IsNull(rs!domingo), "", rs!domingo)
         list_item.SubItems(11) = IIf(IsNull(rs!lunes_fest), "", rs!lunes_fest)
         list_item.SubItems(12) = IIf(IsNull(rs!martes_fest), "", rs!martes_fest)
         list_item.SubItems(13) = IIf(IsNull(rs!miercoles_fest), "", rs!miercoles_fest)
         list_item.SubItems(14) = IIf(IsNull(rs!jueves_fest), "", rs!jueves_fest)
         list_item.SubItems(15) = IIf(IsNull(rs!viernes_fest), "", rs!viernes_fest)
         list_item.SubItems(16) = IIf(IsNull(rs!sabado_fest), "", rs!sabado_fest)
         list_item.SubItems(17) = IIf(IsNull(rs!domingo_fest), "", rs!domingo_fest)
         list_item.SubItems(18) = ""
         list_item.SubItems(19) = 0
         list_item.SubItems(20) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
         var_estatus = IIf(IsNull(rs!lunes_fest), 0, rs!lunes_fest) + IIf(IsNull(rs!martes_fest), 0, rs!martes_fest) + IIf(IsNull(rs!miercoles_fest), 0, rs!miercoles_fest) + IIf(IsNull(rs!jueves_fest), 0, rs!jueves_fest) + IIf(IsNull(rs!viernes_fest), 0, rs!viernes_fest) + IIf(IsNull(rs!sabado_fest), 0, rs!sabado_fest) + IIf(IsNull(rs!domingo_fest), 0, rs!domingo_fest)
         If var_estatus > 0 Then
            list_item.SubItems(19) = 1
         End If
         rs.MoveNext
   Wend
   rs.Close
   For var_i = 1 To Me.lv_clientes.ListItems.Count
       Me.lv_clientes.ListItems.Item(var_i).Selected = True
       If Me.lv_clientes.selectedItem.SubItems(19) = "1" Then
          lv_clientes.ListItems.Item(var_i).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).Bold = True
          lv_clientes.ListItems.Item(var_i).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).ForeColor = &HFF&
       Else
          lv_clientes.ListItems.Item(var_i).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).Bold = False
          lv_clientes.ListItems.Item(var_i).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).ForeColor = &H80000008
       End If
   Next var_i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
       var_si = MsgBox("¿Desea eliminar el dia festivo?", vbYesNo, "ATENCION")
       If var_si = 6 Then
          var_si = MsgBox("Confirmar la eliminacion del dia festivo", vbYesNo, "ATENCION")
          If var_si = 6 Then
             For var_j = 1 To Me.lv_clientes.ListItems.Count
                 Me.lv_clientes.ListItems.Item(var_j).Selected = True
                 If Me.lv_clientes.selectedItem.SubItems(18) = "*" Then
                    rs.Open "UPDATE XXVIA_tB_CLIENTES_RUTAS_DISTR SET LUNES_FEST = NULL, MARTES_fEST = NULL, MIERCOLES_FEST = NULL, JUEVES_FEST = NULL, VIERNES_fEST = NULL, SABADO_FEST = NULL, DOMINGO_FEST = NULL WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                 End If
             Next var_j
          Me.lv_clientes.ListItems.Clear
   'rs.Open "select * from xxvia_vw_dias_despacho  order by nombre_ruta, nombre_Establecimiento", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "select party_site_number, RUTA,NOMBRE_RUTA,a.site_use_id,TITULAR,NOMBRE_TITULAR, LUNES,MARTES,MIERCOLES,JUEVES,VIERNES,SABADO,DOMINGO,PAQUETERIA,LUNES_FEST,MARTES_FEST,MIERCOLES_FEST,JUEVES_FEST,VIERNES_FEST,SABADO_FEST,DOMINGO_FEST, party_site_number||'   '||a.NOMBRE_ESTABLECIMIENTO nombre_establecimiento from xxvia_vw_dias_despacho a, XXVIA_VW_CLIENTES_BCP b where a.site_use_id = to_char(b.site_use_id)  order by nombre_ruta, nombre_Establecimiento ", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_clientes.ListItems.Add(, , rs!nombre_ruta)
         list_item.SubItems(1) = IIf(IsNull(rs!nombre_titular), "", rs!nombre_titular)
         list_item.SubItems(2) = IIf(IsNull(rs!site_use_id), "", rs!site_use_id)
         list_item.SubItems(3) = IIf(IsNull(rs!nombre_Establecimiento), "", rs!nombre_Establecimiento)
         list_item.SubItems(4) = IIf(IsNull(rs!lunes), "", rs!lunes)
         list_item.SubItems(5) = IIf(IsNull(rs!martes), "", rs!martes)
         list_item.SubItems(6) = IIf(IsNull(rs!miercoles), "", rs!miercoles)
         list_item.SubItems(7) = IIf(IsNull(rs!jueves), "", rs!jueves)
         list_item.SubItems(8) = IIf(IsNull(rs!viernes), "", rs!viernes)
         list_item.SubItems(9) = IIf(IsNull(rs!sabado), "", rs!sabado)
         list_item.SubItems(10) = IIf(IsNull(rs!domingo), "", rs!domingo)
         list_item.SubItems(11) = IIf(IsNull(rs!lunes_fest), "", rs!lunes_fest)
         list_item.SubItems(12) = IIf(IsNull(rs!martes_fest), "", rs!martes_fest)
         list_item.SubItems(13) = IIf(IsNull(rs!miercoles_fest), "", rs!miercoles_fest)
         list_item.SubItems(14) = IIf(IsNull(rs!jueves_fest), "", rs!jueves_fest)
         list_item.SubItems(15) = IIf(IsNull(rs!viernes_fest), "", rs!viernes_fest)
         list_item.SubItems(16) = IIf(IsNull(rs!sabado_fest), "", rs!sabado_fest)
         list_item.SubItems(17) = IIf(IsNull(rs!domingo_fest), "", rs!domingo_fest)
         list_item.SubItems(18) = ""
         list_item.SubItems(19) = 0
         list_item.SubItems(20) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
         var_estatus = IIf(IsNull(rs!lunes_fest), 0, rs!lunes_fest) + IIf(IsNull(rs!martes_fest), 0, rs!martes_fest) + IIf(IsNull(rs!miercoles_fest), 0, rs!miercoles_fest) + IIf(IsNull(rs!jueves_fest), 0, rs!jueves_fest) + IIf(IsNull(rs!viernes_fest), 0, rs!viernes_fest) + IIf(IsNull(rs!sabado_fest), 0, rs!sabado_fest) + IIf(IsNull(rs!domingo_fest), 0, rs!domingo_fest)
         If var_estatus > 0 Then
            list_item.SubItems(19) = 1
         End If
         rs.MoveNext
   Wend
   rs.Close
   For var_i = 1 To Me.lv_clientes.ListItems.Count
       Me.lv_clientes.ListItems.Item(var_i).Selected = True
       If Me.lv_clientes.selectedItem.SubItems(19) = "1" Then
          lv_clientes.ListItems.Item(var_i).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).Bold = True
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).Bold = True
          lv_clientes.ListItems.Item(var_i).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).ForeColor = &HFF&
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).ForeColor = &HFF&
       Else
          lv_clientes.ListItems.Item(var_i).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).Bold = False
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).Bold = False
          lv_clientes.ListItems.Item(var_i).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(10).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(11).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(12).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(13).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(14).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(15).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(16).ForeColor = &H80000008
          lv_clientes.ListItems.Item(var_i).ListSubItems(17).ForeColor = &H80000008
       End If
   Next var_i
          
          
          End If
       End If
    End If
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_clientes.ListItems.Count > 0 Then
         i = lv_clientes.selectedItem.Index
         If lv_clientes.selectedItem.SubItems(18) = "*" Then
            lv_clientes.selectedItem.SubItems(18) = ""
            lv_clientes.ListItems.Item(i).Bold = False
            lv_clientes.ListItems.Item(i).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(5).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(6).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(7).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(8).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(9).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(10).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(11).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(12).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(13).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(14).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(15).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(16).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(17).Bold = False
            lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(12).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(13).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(14).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(15).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(16).ForeColor = &H80000012
            lv_clientes.ListItems.Item(i).ListSubItems(17).ForeColor = &H80000012
            lv_clientes.Refresh
         Else
            lv_clientes.selectedItem.SubItems(18) = "*"
            lv_clientes.ListItems.Item(i).Bold = True
            lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(7).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(8).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(9).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(10).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(11).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(12).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(13).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(14).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(15).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(16).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(17).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(12).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(13).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(14).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(15).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(16).ForeColor = &HFF0000
            lv_clientes.ListItems.Item(i).ListSubItems(17).ForeColor = &HFF0000
            lv_clientes.Refresh
         End If
      End If
   End If

End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_clientes, Me.txt_establecimiento, False)
      txt_buscar = ""
      'pro_textos
   End If

End Sub
