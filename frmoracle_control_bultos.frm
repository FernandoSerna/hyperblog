VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_control_bultos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verificador de Kanbans"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_lista 
      Height          =   2880
      Left            =   1020
      TabIndex        =   3
      Top             =   240
      Width           =   9225
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2415
         Left            =   45
         TabIndex        =   4
         Top             =   375
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4260
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   9135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4140
      Left            =   30
      TabIndex        =   7
      Top             =   900
      Width           =   15090
      Begin MSComctlLib.ListView lv_kanbans 
         Height          =   3975
         Left            =   45
         TabIndex        =   2
         Top             =   135
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   7011
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Kanban"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha creación"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha envío"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha entrada"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Código"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Descripcion"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estatus"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Lunes"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Martes"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Miercoles"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Jueves"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Viernes"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Sabado"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Centro de negocios "
      Height          =   840
      Left            =   45
      TabIndex        =   6
      Top             =   60
      Width           =   15075
      Begin VB.TextBox txt_subinventario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox txt_nombre_subinventario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1830
         TabIndex        =   1
         Top             =   270
         Width           =   13185
      End
   End
End
Attribute VB_Name = "frmoracle_control_bultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim VAR_TIPO_LISTA As Integer

Private Sub Form_Load()
   'Top = 1300
   'Left = 0
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_kanbans_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_kanbans, ColumnHeader)
End Sub

Private Sub lv_kanbans_KeyDown(KeyCode As Integer, Shift As Integer)
   If Me.lv_kanbans.ListItems.Count > 0 Then
      If KeyCode = 116 Then
         var_codigo_kanban = Me.lv_kanbans.selectedItem.SubItems(3)
         var_subinventario_kanban = Me.txt_subinventario
         frmoracle_ubicaciones_kanbans.Show 1
      End If
   End If
End Sub

Private Sub lv_kanbans_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_subinventario.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_subinventario = Me.lv_lista.selectedItem
         Me.txt_nombre_subinventario = Me.lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_subinventario.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_nombre_subinventario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_kanbans.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_subinventario_Change()
   Me.lv_kanbans.ListItems.Clear
   Me.txt_nombre_subinventario = ""
End Sub

Private Sub txt_subinventario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      VAR_TIPO_LISTA = 1
      lbl_lista = "Centros de negocios"
      rs.Open "select secondary_inventory_name , description from mtl_secondary_inventories where attribute3 = 'PTO_VTA'AND DESCRIPTION NOT LIKE '%COSTALES%' ORDER BY DESCRIPTION", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!secondary_inventory_name)
            list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_subinventario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_subinventario.SetFocus
   End If
End Sub

Private Sub txt_subinventario_LostFocus()
   If Me.txt_subinventario <> "" Then
      var_cadena = "select secondary_inventory_name , description from mtl_secondary_inventories where attribute3 = 'PTO_VTA'AND DESCRIPTION NOT LIKE '%COSTALES%' and secondary_inventory_name = ? ORDER BY DESCRIPTION"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = var_cadena
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_subinventario)
           .Parameters.Append parametro
      End With
      Set rsaux1 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux1.EOF Then
         Me.lv_kanbans.ListItems.Clear
         Me.txt_nombre_subinventario = rsaux1!Description
          
         var_cadena = "select vcha_caj_caja_id caja, date_caj_fecha fecha, vcha_Caj_staus estatus, segment1 codigo, description descripcion, pedido_almacen from XXVIA_TB_CAJAS_PROD a, xxvia_system_items_b b where vcha_caj_destino = ? and vcha_Art_articulo_id = b.segment1 and b.organization_id = 93 order by date_Caj_fecha"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = var_cadena
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_subinventario)
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = lv_kanbans.ListItems.Add(, , rs!Caja)
                  list_item.SubItems(1) = IIf(IsNull(rs!Fecha), "", rs!Fecha)
                  
                  var_cadena = "select ncm.DATE_NVM_ENVIADA_CDI AS fecha_envio from XXVIA.XXVIA_TB_SENNAL_SRI_DET srd, XXVIA.XXVIA_TB_CAJAS_PROD cpr, XXPOS.XXPOS_TB_NIVELES_CDI_ML ncm Where srd.VC_CONTENEDOR = cpr.VCHA_CAJ_CAJA_ID and srd.VC_KANBAN_ID = ncm.VC_KANBAN_ID and srd.NUM_SEN_ID = ncm.NUM_SEN_ID and cpr.VCHA_CAJ_CAJA_ID = ? and cpr.VCHA_ART_ARTICULO_ID in (select aps.ARTICULO_ID from XXVIA.XXVIA_TB_ARTICULOS_PS_ICG aps where aps.NUMB_TB_TIPO_CALCULO=3) order by ncm.DATE_NVM_FECHA"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = var_cadena
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!Caja)
                       .Parameters.Append parametro
                  End With
                  'rsaux2.Open "select ncm.DATE_NVM_ENVIADA_CDI AS fecha_envio from XXVIA.XXVIA_TB_SENNAL_SRI_DET srd, XXVIA.XXVIA_TB_CAJAS_PROD cpr, XXPOS.XXPOS_TB_NIVELES_CDI_ML ncm Where srd.VC_CONTENEDOR = cpr.VCHA_CAJ_CAJA_ID and srd.VC_KANBAN_ID = ncm.VC_KANBAN_ID and srd.NUM_SEN_ID = ncm.NUM_SEN_ID and cpr.VCHA_CAJ_CAJA_ID = 'CA181215011828' and cpr.VCHA_ART_ARTICULO_ID in (select aps.ARTICULO_ID from XXVIA.XXVIA_TB_ARTICULOS_PS_ICG aps where aps.NUMB_TB_TIPO_CALCULO=3) order by ncm.DATE_NVM_FECHA", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Set rsaux2 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux2.EOF Then
                     list_item.SubItems(2) = IIf(IsNull(rsaux2!fecha_envio), "", rsaux2!fecha_envio)
                  Else
                     list_item.SubItems(2) = ""
                  End If
                  rsaux2.Close
                  
                  var_cadena = "select srid.VC_CONTENEDOR caja_id, caj.VCHA_ART_ARTICULO_ID, caj.VCHA_CAJ_STAUS, srid.VC_KANBAN_ID, srid.NUM_SEN_ID, niv.DATE_NVM_RECEPCION_CDI2 from XXVIA.XXVIA_TB_SENNAL_SRI_DET srid,     XXPOS.XXPOS_TB_NIVELES_CDI_ML niv,     XXVIA.XXVIA_TB_CAJAS_PROD caj where caj.VCHA_ART_ARTICULO_ID in (select art.ARTICULO_ID from XXVIA.XXVIA_TB_ARTICULOS_PS_ICG art where art.NUMB_TB_TIPO_CALCULO = 3) and srid.VC_KANBAN_ID = niv.VC_KANBAN_ID and srid.NUM_SEN_ID = niv.NUM_SEN_ID and srid.VC_CONTENEDOR = caj.VCHA_CAJ_CAJA_ID and niv.VCHA_NVM_CODIGO = caj.VCHA_ART_ARTICULO_ID and srid.VC_CONTENEDOR = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = var_cadena
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!Caja)
                       .Parameters.Append parametro
                  End With
                  'rsaux2.Open "select ncm.DATE_NVM_ENVIADA_CDI AS fecha_envio from XXVIA.XXVIA_TB_SENNAL_SRI_DET srd, XXVIA.XXVIA_TB_CAJAS_PROD cpr, XXPOS.XXPOS_TB_NIVELES_CDI_ML ncm Where srd.VC_CONTENEDOR = cpr.VCHA_CAJ_CAJA_ID and srd.VC_KANBAN_ID = ncm.VC_KANBAN_ID and srd.NUM_SEN_ID = ncm.NUM_SEN_ID and cpr.VCHA_CAJ_CAJA_ID = 'CA181215011828' and cpr.VCHA_ART_ARTICULO_ID in (select aps.ARTICULO_ID from XXVIA.XXVIA_TB_ARTICULOS_PS_ICG aps where aps.NUMB_TB_TIPO_CALCULO=3) order by ncm.DATE_NVM_FECHA", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Set rsaux2 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux2.EOF Then
                     list_item.SubItems(3) = IIf(IsNull(rsaux2!DATE_NVM_RECEPCION_CDI2), "", rsaux2!DATE_NVM_RECEPCION_CDI2)
                  Else
                     list_item.SubItems(3) = ""
                  End If
                  rsaux2.Close
                  
                  
                  list_item.SubItems(4) = IIf(IsNull(rs!CODIGO), "", rs!CODIGO)
                  list_item.SubItems(5) = IIf(IsNull(rs!descripcion), "", rs!descripcion)
                  var_estatus = IIf(IsNull(rs!estatus), "", rs!estatus)
                  If var_estatus = "R" Then
                     var_Nombre_estatus = "En CEDIS"
                  End If
                  If var_estatus = "S" Then
                     var_Nombre_estatus = "Salio del CEDIS"
                  End If
                  If var_estatus = "E" Then
                     var_Nombre_estatus = "Envio al CEDIS"
                  End If
                  If var_estatus = "O" Then
                     var_Nombre_estatus = "En planta"
                  End If
                  If var_estatus = "A" Then
                     var_Nombre_estatus = "Creado"
                  End If
                  list_item.SubItems(6) = var_Nombre_estatus
                  rs.MoveNext
            Wend
            For var_j = 1 To Me.lv_kanbans.ListItems.Count
                Me.lv_kanbans.ListItems(var_j).Selected = True
                rsaux.Open "select * from tb_oracle_ubicaciones_motor_logistico where clave = '" + Me.txt_subinventario + "' and codigo = '" + Me.lv_kanbans.selectedItem.SubItems(4) + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   Me.lv_kanbans.selectedItem.SubItems(7) = IIf(IsNull(rsaux!ubicacion_1), "", rsaux!ubicacion_1)
                   Me.lv_kanbans.selectedItem.SubItems(8) = IIf(IsNull(rsaux!ubicacion_2), "", rsaux!ubicacion_2)
                   Me.lv_kanbans.selectedItem.SubItems(9) = IIf(IsNull(rsaux!ubicacion_3), "", rsaux!ubicacion_3)
                   Me.lv_kanbans.selectedItem.SubItems(10) = IIf(IsNull(rsaux!ubicacion_4), "", rsaux!ubicacion_4)
                   Me.lv_kanbans.selectedItem.SubItems(11) = IIf(IsNull(rsaux!ubicacion_5), "", rsaux!ubicacion_5)
                   Me.lv_kanbans.selectedItem.SubItems(12) = IIf(IsNull(rsaux!ubicacion_6), "", rsaux!ubicacion_6)
                End If
                rsaux.Close
                
            Next var_j
         
         Else
            MsgBox "No existen kanbans para el destino seleccionado", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "El subinventario no existe", vbOKOnly, "ATENCION"
         Me.txt_subinventario = ""
         Me.txt_nombre_subinventario = ""
      End If
      rsaux1.Close
   End If
End Sub
