VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_busqueda_clientes_rutas_distribucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de clientes"
   ClientHeight    =   11655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11655
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14820
      Picture         =   "frmoracle_busqueda_clientes_rutas_distribucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   60
      Picture         =   "frmoracle_busqueda_clientes_rutas_distribucion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   45
      TabIndex        =   5
      Top             =   270
      Width           =   15105
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   30
      TabIndex        =   3
      Top             =   375
      Width           =   15150
      Begin VB.TextBox txt_clave_esb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4890
         TabIndex        =   11
         Top             =   660
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txt_clave_establecimiento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2460
         TabIndex        =   10
         Top             =   645
         Width           =   1755
      End
      Begin VB.TextBox txt_establecimiento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4245
         TabIndex        =   8
         Top             =   645
         Width           =   10830
      End
      Begin VB.TextBox txt_nombre_titular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2460
         TabIndex        =   0
         Top             =   135
         Width           =   12615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Sucursal:"
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
         Left            =   75
         TabIndex        =   9
         Top             =   735
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
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
         Left            =   75
         TabIndex        =   4
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   10020
      Left            =   30
      TabIndex        =   2
      Top             =   1575
      Width           =   15165
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   9795
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   15030
         _ExtentX        =   26511
         _ExtentY        =   17277
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Titular"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre titular"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Establecimiento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre establecimiento"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Dirección"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Activo"
            Object.Width           =   1235
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_busqueda_clientes_rutas_distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   If Me.lv_clientes.ListItems.Count > 0 Then
      var_si = MsgBox("¿Desea agregar los clientes a la ruta " + var_nombre_ruta_distribucion + "?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To Me.lv_clientes.ListItems.Count
             Me.lv_clientes.ListItems.Item(var_j).Selected = True
             If Me.lv_clientes.selectedItem.SubItems(6) = "*" Then
                rsaux.Open "select * from XXVIA_TB_CLIENTES_RUTAS_DISTR where establecimiento = '" + Me.lv_clientes.selectedItem.SubItems(3) + "' and ruta = '" + var_ruta_distribucion + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                If rsaux.EOF Then
                   'rs.Open "INSERT INTO XXVIA_TB_CLIENTES_RUTAS_DISTR (RUTA, TITULAR, NOMBRE_TITULAR, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, DIRECCION, PRIORIDAD) VALUES ('" + var_ruta_distribucion + "','" + Me.lv_clientes.selectedItem + "','" + Me.lv_clientes.selectedItem.SubItems(1) + "','" + Me.lv_clientes.selectedItem.SubItems(2) + "','" + Replace(Me.lv_clientes.selectedItem.SubItems(3), "'", " ") + "','" + Me.lv_clientes.selectedItem.SubItems(4) + "',0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   rs.Open "INSERT INTO XXVIA_TB_CLIENTES_RUTAS_DISTR (RUTA, TITULAR, NOMBRE_TITULAR, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, DIRECCION, PRIORIDAD) VALUES ('" + var_ruta_distribucion + "','" + Me.lv_clientes.selectedItem + "','" + Me.lv_clientes.selectedItem.SubItems(1) + "','" + Me.lv_clientes.selectedItem.SubItems(3) + "','" + Replace(Me.lv_clientes.selectedItem.SubItems(4), "'", " ") + "','" + Me.lv_clientes.selectedItem.SubItems(4) + "',0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   rsaux2.Open "INSERT INTO TB_ORACLE_BITACORA_RUTAS_CLIENTES (ACCION,RUTA,TITULAR,NOMBRE_TITULAR,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,FECHA_ACCION,USUARIO,MAQUINA) VALUES ('INSERTAR','" + var_ruta_distribucion + "', '" + Me.lv_clientes.selectedItem + "','" + Me.lv_clientes.selectedItem.SubItems(1) + "','" + Me.lv_clientes.selectedItem.SubItems(3) + "','" + Replace(Me.lv_clientes.selectedItem.SubItems(4), "'", " ") + "',GETDATE(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic


                   var_asunto = "Se notifica que el establecimiento " + Me.lv_clientes.selectedItem.SubItems(3) + " " + Me.lv_clientes.selectedItem.SubItems(4) + " del titular " + Me.lv_clientes.selectedItem + " " + Me.lv_clientes.selectedItem.SubItems(1) + " a sido agregado a la ruta " + var_ruta_distribucion + " por el usuario " + var_nombre_usuario_global
                   var_cadena = "call xxvia_pk_correo.sp_enviar_email('','fserna@vianney.com.mx','','','Cambio de ruta del establecimiento " + Me.lv_clientes.selectedItem.SubItems(3) + " " + Me.lv_clientes.selectedItem.SubItems(4) + "','" + var_asunto + "','')"
                   rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic




                'Else
                '   rsaux1.Open "select * from XXVIA_TB_RUTAS_DISTRIBUCION where ruta = '" + rsaux!ruta + "'", cnn, adOpenDynamic, adLockOptimistic
                '   If Not rsaux1.EOF Then
                '      MsgBox "El establecimiento '" + Me.lv_clientes.selectedItem.SubItems(3) + "' esta asignado a la ruta " + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta)
                '   Else
                '      MsgBox "El establecimiento '" + Me.lv_clientes.selectedItem.SubItems(3) + "' esta asignado a otra ruta"
                '   End If
                '   rsaux1.Close
                End If
                rsaux.Close
             End If
         Next var_j
         MsgBox "Se a terminado el proceso de carga de rutas", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If Me.lv_clientes.ListItems.Count > 0 Then
      If KeyAscii = 13 Then
         'rs.Open "select * from XXVIA_TB_CLIENTES_RUTAS_DISTR where ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "'", cnn, adOpenDynamic, adLockOptimistic
         'If Not rs.EOF Then
            'If Me.lv_clientes.selectedItem.SubItems(5) = "" Then
            '   rsaux.Open "select * from XXVIA_TB_RUTAS_DISTRIBUCION where ruta = '" + rs!ruta + "'", cnn, adOpenDynamic, adLockOptimistic
            '   If Not rsaux.EOF Then
            '      MsgBox "El cliente ya fue asignado a la ruta " + rsaux!nombre_ruta, vbOKOnly, "ATENCION"
            '   Else
            '      MsgBox "El cliente ya fue asignado a otra ruta", vbOKOnly, "ATENCION"
            '   End If
            '   rsaux.Close
            'Else
         '      i = lv_clientes.selectedItem.Index
         '      lv_clientes.selectedItem.SubItems(5) = ""
         '      lv_clientes.ListItems.Item(i).Bold = False
         '      'lv_clientes.ListItems.Item(i).ForeColor = &H80000012
         '      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
         '      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
         '      lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
         '      lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
         '      lv_clientes.ListItems.Item(i).ListSubItems(5).Bold = False
         '      lv_clientes.Refresh
         '
         '   End If
         'Else
            i = lv_clientes.selectedItem.Index
            If lv_clientes.selectedItem.SubItems(6) = "*" Then
               lv_clientes.selectedItem.SubItems(6) = ""
               lv_clientes.ListItems.Item(i).Bold = False
               'lv_clientes.ListItems.Item(i).ForeColor = &H80000012
               lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
               lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
               lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
               lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
               lv_clientes.ListItems.Item(i).ListSubItems(5).Bold = False
               lv_clientes.Refresh
            Else
               lv_clientes.selectedItem.SubItems(6) = "*"
               lv_clientes.ListItems.Item(i).Bold = True
               'lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
               lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
               lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
               lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
               lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
               lv_clientes.ListItems.Item(i).ListSubItems(5).Bold = True
               lv_clientes.Refresh
            End If
         
         
         
         'End If
         'rs.Close
      End If
   End If
End Sub

Private Sub Text1_Change()
   Me.lv_clientes.ListItems.Clear
End Sub

Private Sub Text2_Change()
   Me.lv_clientes.ListItems.Clear
End Sub

Private Sub txt_clave_esb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "SELECT SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE  FROM XXVIA_VW_CLIENTES_BCP WHERE  SITE_USE_ID = " + Me.txt_clave_esb + " and site_use_code = 'SHIP_TO'  and attribute3 = 'REV'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_establecimiento = IIf(IsNull(rs!RAZON_SOCIAL_CLIENTE), "", rs!RAZON_SOCIAL_CLIENTE)
         Me.txt_establecimiento.SetFocus
      Else
         MsgBox "Clave de establecimiento incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_clave_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "SELECT SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE  FROM XXVIA_VW_CLIENTES_BCP WHERE  party_site_number = '" + Me.txt_clave_establecimiento + "' and site_use_code = 'SHIP_TO'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_establecimiento = IIf(IsNull(rs!RAZON_SOCIAL_CLIENTE), "", rs!RAZON_SOCIAL_CLIENTE)
         Me.txt_establecimiento.SetFocus
      Else
         MsgBox "Clave de establecimiento incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_establecimiento_Change()
   Me.lv_clientes.ListItems.Clear
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.txt_nombre_titular <> "" Then
         If Mid(Me.txt_nombre_titular, 1, 3) = "VTH" Then
            var_cadena = Mid(Me.txt_nombre_titular, 5, Len(Me.txt_nombre_titular))
            rs.Open "select secondary_inventory_name VCHA_ALM_ALMACEN_ID, description VCHA_ALM_NOMBRE  from mtl_secondary_inventories where ATTRIBUTE3 LIKE '%PTO%' AND ORGANIZATION_ID = 93 and description like '%" + var_cadena + "%'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_clientes.ListItems.Add(, , 2040)
                  list_item.SubItems(1) = "VIANNEY TEXTIL HOGAR S.A DE C.V."
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                  list_item.SubItems(3) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                  list_item.SubItems(4) = Format(rs!vcha_alm_nombre)
                  list_item.SubItems(5) = ""
                  list_item.SubItems(6) = "REV"
                  rs.MoveNext
            Wend
            rs.Close
         Else
            If Me.txt_nombre_titular <> "" Then
               rs.Open "SELECT CUST_ACCOUNT_ID , ACCOUNT_FULL_NAME NOMBRE_TITULAR ,SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, PARTY_SITE_NUMBER,CALLE||' '||COLONIA||' '||' '||CIUDAD AS DIRECCION,attribute3 FROM XXVIA_VW_CLIENTES_BCP WHERE ACCOUNT_FULL_NAME LIKE '%" + Me.txt_nombre_titular + "%' and RAZON_SOCIAL_CLIENTE like '%" + Me.txt_establecimiento + "%' and site_use_code = 'SHIP_TO'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               rs.Open "SELECT CUST_ACCOUNT_ID , ACCOUNT_FULL_NAME NOMBRE_TITULAR ,SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, PARTY_SITE_NUMBER,CALLE||' '||COLONIA||' '||' '||CIUDAD AS DIRECCION, attribute3 FROM XXVIA_VW_CLIENTES_BCP WHERE RAZON_SOCIAL_CLIENTE like '%" + Me.txt_establecimiento + "%' and site_use_code = 'SHIP_TO'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            While Not rs.EOF
                  Set list_item = Me.lv_clientes.ListItems.Add(, , rs!CUST_ACCOUNT_ID)
                  list_item.SubItems(1) = Format(rs!nombre_titular)
                  list_item.SubItems(2) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
                  list_item.SubItems(3) = IIf(IsNull(rs!CLAVE), "", rs!CLAVE)
                  list_item.SubItems(4) = Format(rs!nombre_Establecimiento)
                  list_item.SubItems(5) = IIf(IsNull(rs!direccion), "", rs!direccion)
                  list_item.SubItems(6) = IIf(IsNull(rs!attribute3), "", rs!attribute3)
                  rs.MoveNext
            Wend
            rs.Close
         End If
      Else
         rs.Open "SELECT CUST_ACCOUNT_ID , ACCOUNT_FULL_NAME NOMBRE_TITULAR ,SITE_USE_ID CLAVE, PARTY_SITE_NUMBER,RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, CALLE||' '||COLONIA||' '||' '||CIUDAD AS DIRECCION, attribute3 FROM XXVIA_VW_CLIENTES_BCP WHERE RAZON_SOCIAL_CLIENTE like '%" + Me.txt_establecimiento + "%' and site_use_code = 'SHIP_TO'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_clientes.ListItems.Add(, , rs!CUST_ACCOUNT_ID)
               list_item.SubItems(1) = Format(rs!nombre_titular)
               list_item.SubItems(2) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
               list_item.SubItems(3) = IIf(IsNull(rs!CLAVE), "", rs!CLAVE)
               list_item.SubItems(4) = Format(rs!nombre_Establecimiento)
               list_item.SubItems(5) = IIf(IsNull(rs!direccion), "", rs!direccion)
               list_item.SubItems(6) = IIf(IsNull(rs!attribute3), "", rs!attribute3)
               
               rs.MoveNext
         Wend
         rs.Close
      
      End If
   End If
End Sub

Private Sub txt_nombre_titular_Change()
   Me.lv_clientes.ListItems.Clear
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.txt_nombre_titular <> "" Then
         If Mid(Me.txt_nombre_titular, 1, 3) = "VTH" Then
            var_cadena = Mid(Me.txt_nombre_titular, 5, Len(Me.txt_nombre_titular))
            rs.Open "select secondary_inventory_name VCHA_ALM_ALMACEN_ID, description VCHA_ALM_NOMBRE  from mtl_secondary_inventories where ATTRIBUTE3 LIKE '%PTO%' AND ORGANIZATION_ID = 93 and description like '%" + var_cadena + "%'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_clientes.ListItems.Add(, , 2040)
                  list_item.SubItems(1) = "VIANNEY TEXTIL HOGAR S.A DE C.V."
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                  list_item.SubItems(3) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                  list_item.SubItems(4) = Format(rs!vcha_alm_nombre)
                  list_item.SubItems(5) = ""
                  rs.MoveNext
            Wend
            rs.Close
         Else
            'rs.Open "SELECT CUST_ACCOUNT_ID , ACCOUNT_FULL_NAME NOMBRE_TITULAR ,SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, CALLE||' '||COLONIA||' '||' '||CIUDAD AS DIRECCION FROM XXVIA_VW_CLIENTES_BCP WHERE ACCOUNT_FULL_NAME LIKE '%" + Me.txt_nombre_titular + "%' AND SITE_USE_CODE = 'SHIP_TO'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "SELECT CUST_ACCOUNT_ID , PARTY_SITE_NUMBER,ACCOUNT_FULL_NAME NOMBRE_TITULAR ,SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, CALLE||' '||COLONIA||' '||' '||CIUDAD AS DIRECCION FROM XXVIA_VW_CLIENTES_BCP WHERE ACCOUNT_FULL_NAME LIKE '%" + Me.txt_nombre_titular + "%'  and site_use_code = 'SHIP_TO' and attribute3 = 'REV'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_clientes.ListItems.Add(, , rs!CUST_ACCOUNT_ID)
                  list_item.SubItems(1) = Format(rs!nombre_titular)
                  list_item.SubItems(2) = IIf(IsNull(rs!party_site_number), "", rs!party_site_number)
                  list_item.SubItems(3) = IIf(IsNull(rs!CLAVE), "", rs!CLAVE)
                  list_item.SubItems(4) = Format(rs!nombre_Establecimiento)
                  list_item.SubItems(5) = IIf(IsNull(rs!direccion), "", rs!direccion)
                  rs.MoveNext
            Wend
            rs.Close
         End If
      End If
   End If
End Sub
