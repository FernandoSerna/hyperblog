VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frminformacion_articulos_enviar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información de artículos a enviar"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6015
   Begin VB.Frame Frame2 
      Caption         =   " Artículos "
      Height          =   3990
      Left            =   90
      TabIndex        =   4
      Top             =   375
      Width           =   5865
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frminformacion_articulos_enviar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   75
         Picture         =   "frminformacion_articulos_enviar.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Picture         =   "frminformacion_articulos_enviar.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         Picture         =   "frminformacion_articulos_enviar.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar (Enter)"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1395
         Picture         =   "frminformacion_articulos_enviar.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   285
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   3255
         Left            =   60
         TabIndex        =   10
         Top             =   630
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   5741
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
            Text            =   "Clave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6932
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Email"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   150
      TabIndex        =   3
      Top             =   330
      Width           =   5805
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5625
      Picture         =   "frminformacion_articulos_enviar.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frminformacion_articulos_enviar.frx":0E84
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Enviar Información"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frminformacion_articulos_enviar.frx":0F86
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2475
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1545
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frminformacion_articulos_enviar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function


Private Sub cmd_correo_Click()
   Dim var_numero_clientes As Integer
   Dim var_correo_electronico As String
   var_eliminar = DeleteFile(var_ruta & "estilo.dbf")
   var_copia = CopyFile(var_ruta & "testilo.dbf", var_ruta & "estilo.dbf", 1)
   
   var_eliminar = DeleteFile(var_ruta & "desdesco.dbf")
   var_copia = CopyFile(var_ruta & "tdesdesco.dbf", var_ruta & "desdesco.dbf", 1)
   
   var_eliminar = DeleteFile(var_ruta & "catalogo.dbf")
   var_copia = CopyFile(var_ruta & "tcatalogo.dbf", var_ruta & "catalogo.dbf", 1)
   rs.Open "SELECT     dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, dbo.TB_Articulos.MONE_ART_COSTO_ESTANDAR , dbo.TB_LINEAS.VCHA_LIN_LINEA_ID, dbo.TB_LINEAS.VCHA_LIN_NOMBRE FROM         dbo.TB_ARTICULOS LEFT OUTER JOIN dbo.TB_LINEAS ON dbo.TB_ARTICULOS.VCHA_LIN_LINEA_ID = dbo.TB_LINEAS.VCHA_LIN_LINEA_ID ", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            var_cadena = "insert into estilo (cveestilo, descripcio, costo, preciolist, tallas, talla1, talla2, talla3, talla4, talla5, talla6, linea, nombre) values "
            var_cadena = var_cadena + "('" + IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id) + "', '" + IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español) + "', " + CStr(IIf(IsNull(rs!MONE_ART_COSTO_ESTANDAR), 0, rs!MONE_ART_COSTO_ESTANDAR)) + ", " + CStr(IIf(IsNull(rs!MONE_ART_PRECIO_BASE), 0, rs!MONE_ART_PRECIO_BASE)) + ", 0,0,0,0,0,0,0, '" + "','" + "')"
            Text1 = var_cadena
            rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   rs.Open "select distinct vcha_can_canal_venta_id, vcha_cat_catalogo_id, vcha_cat_nombre, dtim_vig_fecha_fin from VW_CATALOGOS_CANALES_VENTA", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            If IsNull(rs!dtim_vig_fecha_fin) Then
               var_tipo = "T"
            Else
               If rs!dtim_vig_fecha_fin < Date Then
                  var_tipo = "F"
               Else
                  var_tipo = "T"
               End If
            End If
            var_cadena = "insert into catalogo (canalvta, cvecatalog, descripcio, ano, vigente, fecha_venc) values ('" + IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id) + "', '" + IIf(IsNull(rs!vcha_Cat_catalogo_id), "", rs!vcha_Cat_catalogo_id) + "', '" + IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre) + "', 0, ." + var_tipo + "., ctod('" + CStr(IIf(IsNull(rs!dtim_vig_fecha_fin), "", rs!dtim_vig_fecha_fin)) + "'))"
            rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
   rs.Open "select * from TB_DESCUENTOS_CATALOGOS", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            rsaux.Open "insert into desdesco (canalvta, liminf, limsup, dtopesos, dtodollar) values ('" + IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id) + "', " + CStr(IIf(IsNull(rs!INTE_DES_LIMITE_INFERIOR), 0, rs!INTE_DES_LIMITE_INFERIOR)) + ", " + CStr(IIf(IsNull(rs!INTE_DES_LIMITE_SUPERIOR), "", rs!INTE_DES_LIMITE_SUPERIOR)) + ", " + CStr(IIf(IsNull(rs!FLOA_DES_DESCUENTO), 0, rs!FLOA_DES_DESCUENTO)) + ",0)", var_tabla, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
   var_numero_clientes = lv_clientes.ListItems.Count
   var_correo_electronico = ""
   If var_numero_clientes > 0 Then
      For var_i = 1 To var_numero_clientes
          lv_clientes.ListItems.Item(var_i).Selected = True
          If lv_clientes.selectedItem.SubItems(2) = "*" Then
             If Trim(lv_clientes.selectedItem.SubItems(3)) <> "" Then
                If Trim(var_correo_electronico) <> "" Then
                   var_correo_electronico = var_correo_electronico + ";" + lv_clientes.selectedItem.SubItems(3)
                Else
                   var_correo_electronico = lv_clientes.selectedItem.SubItems(3)
                End If
             End If
          End If
      Next var_i
   End If
   If Trim(var_correo_electronico) <> "" Then
      If MAPISession1.SessionID = 0 Then
         MAPISession1.SignOn
      End If
      MAPIMessages1.SessionID = MAPISession1.SessionID
      MAPIMessages1.Compose
      MAPIMessages1.RecipDisplayName = var_correo_electronico
      MAPIMessages1.RecipAddress = var_correo_electronico
      MAPIMessages1.AddressResolveUI = True
      MAPIMessages1.ResolveName
      MAPIMessages1.MsgSubject = "Archivos de catalogos de articulos"
      MAPIMessages1.MsgNoteText = "Se adjunta archivos de articulos"
      MAPIMessages1.AttachmentPathName = var_ruta + "tclientes.dbf"
      MAPIMessages1.Send False
      If MAPISession1.SessionID > 0 Then
         MAPISession1.SignOff
      End If
   End If
End Sub

Private Sub cmd_invertir_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_clientes.ListItems.Count
   For i = 1 To n
       If lv_clientes.ListItems.Item(i).SubItems(2) = "*" Then
          lv_clientes.ListItems.Item(i).SubItems(2) = " "
          lv_clientes.ListItems.Item(i).Bold = False
          lv_clientes.ListItems.Item(i).ForeColor = &H80000012
          lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_clientes.ListItems.Item(i).SubItems(2) = "*"
          lv_clientes.ListItems.Item(i).Bold = True
          lv_clientes.ListItems.Item(i).ForeColor = &H8000&
          lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
       End If
   Next
   lv_clientes.Refresh
End Sub

Private Sub cmd_marcar_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = lv_clientes.selectedItem.Index
   lv_clientes.ListItems.Item(i).SubItems(2) = "*"
   lv_clientes.ListItems.Item(i).Bold = True
   lv_clientes.ListItems.Item(i).ForeColor = &H8000&
   lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
   lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
   lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
   lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   lv_clientes.Refresh
End Sub

Private Sub cmd_ninguno_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_clientes.ListItems.Count
   For i = 1 To n
       lv_clientes.ListItems.Item(i).SubItems(2) = " "
       lv_clientes.ListItems.Item(i).Bold = False
       lv_clientes.ListItems.Item(i).ForeColor = &H80000012
       lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next
   lv_clientes.Refresh

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
         primera_vez = False
         segunda_vez = False
         n = lv_clientes.ListItems.Count
         For i = 1 To n
            If lv_clientes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
               numero_seleccionado1 = i
               primera_vez = True
            End If
            If lv_clientes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
               numero_seleccionado2 = i
            End If
         Next
         For i = numero_seleccionado1 To numero_seleccionado2
            lv_clientes.ListItems.Item(i).SubItems(2) = "*"
            lv_clientes.ListItems.Item(i).Bold = True
            lv_clientes.ListItems.Item(i).ForeColor = &H8000&
            lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
            lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
            lv_clientes.Refresh
         Next


End Sub

Private Sub cmd_todos_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_clientes.ListItems.Count
   For i = 1 To n
       lv_clientes.ListItems.Item(i).SubItems(2) = "*"
       lv_clientes.ListItems.Item(i).Bold = True
       lv_clientes.ListItems.Item(i).ForeColor = &H8000&
       lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_clientes.Refresh
End Sub

Private Sub Command1_Click()
   Dim lon As Integer
   Dim i As Integer
   Dim var_nombre As String
   Dim var_nombre_nuevo As String
   Dim var_clave As String
   rs.Open "select * from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      lon = Len(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
      var_clave = rs!vcha_art_articulo_id
      var_nombre = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
      var_nombre_nuevo = ""
      For i = 1 To lon
          If Mid(var_nombre, i, 1) <> "'" Then
             var_nombre_nuevo = var_nombre_nuevo + Mid(var_nombre, i, 1)
          End If
      Next i
      rsaux.Open "update tb_Articulos set vcha_art_nombre_español = '" + var_nombre_nuevo + "' where vcha_Art_articulo_id ='" + var_clave + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.MoveNext
   Wend
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1500
   Left = 2800
   Set var_tabla = CreateObject("ADODB.connection")
   Dim list_item As ListItem
   rs.Open "select * from vw_clientes_1 where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_clientes.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            list_item.SubItems(2) = ""
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_CLI_EMAIL), "", rs!VCHA_CLI_EMAIL)
            rs.MoveNext
      Wend
   End If
   rs.Close
   rs.Open "select VCHA_PRI_RUTA_ARCHIVOS_ARTICULOS from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   var_ruta = IIf(IsNull(rs(0).Value), "", rs(0).Value)
   rs.Close
   
   If Trim(var_ruta) <> "" Then
      var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   Else
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_informacion_articulos_enviar)
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub
