VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmimportacion_codigos_SIP_SID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacón de códigos SIP SID"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3630
      Picture         =   "frmimportacion_codigos_SIP_SID.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Buscar artículos"
      Top             =   765
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmimportacion_codigos_SIP_SID.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Importar"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7335
      Picture         =   "frmimportacion_codigos_SIP_SID.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   30
      TabIndex        =   1
      Top             =   270
      Width           =   7680
   End
   Begin VB.Frame x 
      Height          =   5625
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   7650
      Begin VB.Frame frm_mensaje 
         BackColor       =   &H8000000E&
         Height          =   690
         Left            =   1815
         TabIndex        =   8
         Top             =   2310
         Width           =   4290
         Begin VB.Label lbl_mensaje 
            BackColor       =   &H8000000E&
            Caption         =   "Se estan buscando los códigos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   225
            TabIndex        =   9
            Top             =   210
            Width           =   3900
         End
      End
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
         Height          =   450
         Left            =   1920
         TabIndex        =   6
         Top             =   195
         Width           =   1620
      End
      Begin MSComctlLib.ListView lv_codigos 
         Height          =   4890
         Left            =   45
         TabIndex        =   4
         Top             =   675
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   8625
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Planta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Alta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de importación:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   1575
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   15
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
      Left            =   810
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmimportacion_codigos_SIP_SID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim cnn_cantia As ADODB.Connection

Private Sub cmd_eliminar_Click()
   Dim sum1 As Integer
   Dim sum2 As Integer
   Dim icont As Integer
   Dim VERIFICADOR As Integer
   Dim verificador2 As Integer
   Dim var_codigo As String
   Dim longitud As Integer
   Dim msuma As Integer
   Me.lv_codigos.ListItems.Clear
   Me.frm_mensaje.Visible = True
   If var_empresa = "15" Then
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID in ('17')", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
            If var_conexion_cantia <> "" Then
               cnn_cantia.Open var_conexion_cantia
               var_cadena = "select vcha_pro_producto_id, vcha_pro_descripcionespañol, dtim_pro_Fechaalta, vcha_aud_usuario, mon_pro_preciolista as mon_pro_precio_lista  from tb_producto"
               rsaux.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     txt_codigo = Trim(CStr(rsaux!vcha_pro_producto_id))
                     rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux2.EOF Then
                        Set list_item = Me.lv_codigos.ListItems.Add(, , txt_codigo)
                        list_item.SubItems(1) = UCase(IIf(IsNull(rsaux!vcha_pro_descripcionespañol), "", rsaux!vcha_pro_descripcionespañol))
                        list_item.SubItems(2) = rs!VCHA_UOR_NOMBRE
                        list_item.SubItems(3) = IIf(IsNull(rsaux!dtim_pro_Fechaalta), "", rsaux!dtim_pro_Fechaalta)
                        VAR_FECHA_SIP = IIf(IsNull(rsaux!dtim_pro_Fechaalta), "", rsaux!dtim_pro_Fechaalta)
                        var_dia = CStr(Day(VAR_FECHA_SIP))
                        var_mes = CStr(Month(VAR_FECHA_SIP))
                        var_año = CStr(Year(VAR_FECHA_SIP))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha_sip_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        var_dia = CStr(Day(Date))
                        var_mes = CStr(Month(Date))
                        var_año = CStr(Year(Date))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        VAR_FECHA_STR = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        list_item.SubItems(4) = IIf(IsNull(rsaux!vcha_aud_usuario), "", rsaux!vcha_aud_usuario)
                        list_item.SubItems(5) = IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista)
                        list_item.SubItems(6) = IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista) * 0.65
                        'MsgBox "INSERT INTO TB_aRTICULOS (VCHA_aRT_ARTICULO_ID, VCHA_aRT_NOMBRE_eSPAÑOL, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_INSERCION, VCHA_ART_USUARIO_INSERCION, DTIM_ART_FECHA_ALTA_SIP, VCHA_ART_PLANTA_ORIGEN, mone_Art_precio_base, mone_Art_costo_Estandar, vcha_emp_Emrpesa_id) VALUES ('" + txt_codigo + "','" + UCase(IIf(IsNull(rsaux!vcha_pro_descripcionespañol), "", rsaux!vcha_pro_descripcionespañol)) + "'," + VAR_FECHA_STR + "," + VAR_FECHA_STR + ",'" + IIf(IsNull(rsaux!vcha_aud_usuario), "", rsaux!vcha_aud_usuario) + "'," + var_fecha_sip_str + ",'" + rs!VCHA_UOR_NOMBRE + "'," + CStr(IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista)) + ", " + CStr(IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista) * 0.65) + ")"
                        rsaux3.Open "INSERT INTO TB_aRTICULOS (VCHA_aRT_ARTICULO_ID, VCHA_aRT_NOMBRE_eSPAÑOL, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_INSERCION, VCHA_ART_USUARIO_INSERCION, DTIM_ART_FECHA_ALTA_SIP, VCHA_ART_PLANTA_ORIGEN, mone_Art_precio_base, mone_Art_costo_Estandar, vcha_emp_EmPRESA_id) VALUES ('" + txt_codigo + "','" + Mid(UCase(IIf(IsNull(rsaux!vcha_pro_descripcionespañol), "", rsaux!vcha_pro_descripcionespañol)), 1, 50) + "'," + VAR_FECHA_STR + "," + VAR_FECHA_STR + ",'" + IIf(IsNull(rsaux!vcha_aud_usuario), "", rsaux!vcha_aud_usuario) + "'," + var_fecha_sip_str + ",'" + rs!VCHA_UOR_NOMBRE + "'," + CStr(IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista)) + ", " + CStr(IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista) * 0.65) + ", '" + var_empresa + "')", cnn, adOpenDynamic, adLockOptimistic
                        If IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista) > 1 Then
                           rsaux3.Open "iNSert into tb_Detalle_lista_precios (vcha_lis_lista_precios_id, vcha_Art_articulo_id, floa_dli_precio) values ('01','" + txt_codigo + "'," + CStr(IIf(IsNull(rsaux!mon_pro_precio_lista), 0, rsaux!mon_pro_precio_lista)) + ") ", cnn, adOpenDynamic, adLockOptimistic
                        End If
                     End If
                     rsaux2.Close
                     
                     rsaux.MoveNext
                     Me.frm_mensaje.Refresh
                     Me.lv_codigos.Refresh
               Wend
               rsaux.Close
               cnn_cantia.Close
            End If
            rs.MoveNext
      Wend
      rs.Close
   Else
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID in ('01','02','03','04')", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
            If var_conexion_cantia <> "" Then
               cnn_cantia.Open var_conexion_cantia
               var_cadena = "select vcha_pro_producto_id, vcha_pro_descripcionespañol, dtim_pro_Fechaalta, vcha_aud_usuario from tb_producto where len(vcha_pro_producto_id) = 5 and dtim_pro_fechaalta >= {d '2009-03-01'} and (vcha_pro_producto_id not LIKE '%A%' AND vcha_pro_producto_id not LIKE '%B%' AND vcha_pro_producto_id not LIKE '%C%' AND vcha_pro_producto_id not LIKE '%D%' AND vcha_pro_producto_id not LIKE '%E%' AND vcha_pro_producto_id not LIKE '%F%' AND vcha_pro_producto_id not LIKE '%G%' AND vcha_pro_producto_id not LIKE '%H%' AND vcha_pro_producto_id not LIKE '%I%' AND vcha_pro_producto_id not LIKE '%J%' AND vcha_pro_producto_id not LIKE '%K%' AND vcha_pro_producto_id not LIKE '%L%' AND vcha_pro_producto_id not LIKE '%M%' AND vcha_pro_producto_id not LIKE '%N%' AND vcha_pro_producto_id not LIKE '%O%' AND vcha_pro_producto_id not LIKE '%P%' AND vcha_pro_producto_id not LIKE '%Q%' AND vcha_pro_producto_id not LIKE '%R%' AND vcha_pro_producto_id not LIKE '%S%' AND "
               var_cadena = var_cadena + " vcha_pro_producto_id not LIKE '%T%' AND vcha_pro_producto_id not LIKE '%U%' AND vcha_pro_producto_id not"
               var_cadena = var_cadena + " LIKE '%W%' AND vcha_pro_producto_id not LIKE '%X%' AND vcha_pro_producto_id not LIKE '%Y%' AND vcha_pro_producto_id not LIKE '%Z%')"
               rsaux.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     txt_codigo = "646244" + Trim(CStr(rsaux!vcha_pro_producto_id))
                     sum1 = 0
                     sum2 = 0
                     mcodigo = txt_codigo
                     longitud = Len(mcodigo)
                     For icont = 1 To longitud
                         If ((icont / 2) - Int((icont / 2))) = 0 Then
                            sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                         Else
                            sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                         End If
                     Next icont
                     msuma = sum1 * 13 + sum2
                     VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                     If VERIFICADOR = 10 Then
                        VERIFICADOR = 0
                     End If
                     txt_codigo = txt_codigo + Trim(CStr(VERIFICADOR))
                     rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rsaux2.EOF Then
                        Set list_item = Me.lv_codigos.ListItems.Add(, , txt_codigo)
                        list_item.SubItems(1) = UCase(IIf(IsNull(rsaux!vcha_pro_descripcionespañol), "", rsaux!vcha_pro_descripcionespañol))
                        list_item.SubItems(2) = rs!VCHA_UOR_NOMBRE
                        list_item.SubItems(3) = IIf(IsNull(rsaux!dtim_pro_Fechaalta), "", rsaux!dtim_pro_Fechaalta)
                        VAR_FECHA_SIP = IIf(IsNull(rsaux!dtim_pro_Fechaalta), "", rsaux!dtim_pro_Fechaalta)
                        var_dia = CStr(Day(VAR_FECHA_SIP))
                        var_mes = CStr(Month(VAR_FECHA_SIP))
                        var_año = CStr(Year(VAR_FECHA_SIP))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha_sip_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        var_dia = CStr(Day(Date))
                        var_mes = CStr(Month(Date))
                        var_año = CStr(Year(Date))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        VAR_FECHA_STR = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        list_item.SubItems(4) = IIf(IsNull(rsaux!vcha_aud_usuario), "", rsaux!vcha_aud_usuario)
                        rsaux3.Open "INSERT INTO TB_aRTICULOS (VCHA_aRT_ARTICULO_ID, VCHA_aRT_NOMBRE_eSPAÑOL, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_INSERCION, VCHA_ART_USUARIO_INSERCION, DTIM_ART_FECHA_ALTA_SIP, VCHA_ART_PLANTA_ORIGEN) VALUES ('" + txt_codigo + "','" + UCase(IIf(IsNull(rsaux!vcha_pro_descripcionespañol), "", rsaux!vcha_pro_descripcionespañol)) + "'," + VAR_FECHA_STR + "," + VAR_FECHA_STR + ",'" + IIf(IsNull(rsaux!vcha_aud_usuario), "", rsaux!vcha_aud_usuario) + "'," + var_fecha_sip_str + ",'" + rs!VCHA_UOR_NOMBRE + "')", cnn, adOpenDynamic, adLockOptimistic
                        
                     End If
                     rsaux2.Close
                     
                     rsaux.MoveNext
                     Me.frm_mensaje.Refresh
                     Me.lv_codigos.Refresh
               Wend
               rsaux.Close
               cnn_cantia.Close
            End If
            rs.MoveNext
      Wend
      rs.Close
   End If
   Me.frm_mensaje.Visible = False
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If IsDate(Me.txt_fecha) Then
      var_dia = CStr(Day(CDate(Me.txt_fecha)))
      var_mes = CStr(Month(CDate(Me.txt_fecha)))
      var_año = CStr(Year(CDate(Me.txt_fecha)))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      VAR_FECHA_STR = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      var_fecha_crystal = var_año + "," + var_mes + "," + var_dia
      rs.Open "SELECT * FROM TB_aRTICULOS WHERE DTIM_ART_FECHA_INSERCION = " + VAR_FECHA_STR, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set reporte = appl.OpenReport(App.Path + "\rep_articulos_insertados_desde_sip.rpt")
         reporte.RecordSelectionFormula = " {VW_ARTICULOS_INSERTADAS_DESDE_SIP.DTIM_ART_FECHA_INSERCION} = date(" + var_fecha_crystal + ")"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\Articulos_insertados_desde_SIP_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         var_correo_electronico = "fserna@vianney.com.mx"
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
            MAPIMessages1.MsgSubject = "Artículos insertados al SID desde el SIP con fecha de " + Me.txt_fecha
            MAPIMessages1.MsgNoteText = "Se adjunta archivo con artículos insertados desde el SIP con fecha de " + Me.txt_fecha
            MAPIMessages1.AttachmentIndex = 0
            MAPIMessages1.AttachmentPathName = archivo
            'MAPIMessages1.AttachmentIndex = 1
            'MAPIMessages1.AttachmentPathName = archivo
                                
                                
            MAPIMessages1.Send False
            If MAPISession1.SessionID > 0 Then
               MAPISession1.SignOff
            End If
         Else
            MsgBox "El usuario no cuenta con cuenta de correo", vbOKOnly, "ATENCION"
         End If
         
         
         
      Else
         MsgBox "No existen artículos insertados en la fecha seleccionada", vbOKOnly, "ATENCION"
      End If
      If rs.State = 1 Then
         rs.Close
      End If
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Me.frm_mensaje.Visible = False
   Set cnn_cantia = CreateObject("ADODB.connection")
   Top = 800
   Left = 1900
   Me.txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_codigos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_codigos, ColumnHeader)
End Sub
