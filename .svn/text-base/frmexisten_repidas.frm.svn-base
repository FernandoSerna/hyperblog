VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmexisten_rapidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda Rapida de Existencias"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   Icon            =   "frmexisten_repidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_ubicaciones 
      Height          =   4785
      Left            =   2595
      TabIndex        =   25
      Top             =   1545
      Width           =   6735
      Begin MSComctlLib.ListView lv_ubicaciones 
         Height          =   4125
         Left            =   60
         TabIndex        =   26
         Top             =   600
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   7276
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
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl_ubicacion 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   6660
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Height          =   330
      Left            =   135
      Picture         =   "frmexisten_repidas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6585
      Width           =   345
   End
   Begin VB.Frame Frame5 
      Caption         =   " Ubicaciones "
      Height          =   750
      Left            =   60
      TabIndex        =   17
      Top             =   2325
      Width           =   11505
      Begin VB.TextBox txt_ubicacion_1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   1395
         TabIndex        =   20
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   5325
         TabIndex        =   19
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   8925
         TabIndex        =   18
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 1:"
         Height          =   195
         Left            =   450
         TabIndex        =   23
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 2:"
         Height          =   195
         Left            =   4395
         TabIndex        =   22
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 3:"
         Height          =   195
         Left            =   7995
         TabIndex        =   21
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.TextBox txt_cantidad_ordenes 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9255
      TabIndex        =   16
      Top             =   6540
      Width           =   2265
   End
   Begin VB.TextBox txt_nombre_almacen 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2085
      TabIndex        =   2
      Top             =   315
      Width           =   9300
   End
   Begin VB.Frame Frame4 
      Caption         =   " Almacen "
      Height          =   690
      Left            =   75
      TabIndex        =   12
      Top             =   60
      Width           =   11475
      Begin VB.TextBox txt_clave_almacen 
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   255
         Width           =   1890
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Desgloce "
      Height          =   3345
      Left            =   75
      TabIndex        =   11
      Top             =   3075
      Width           =   11490
      Begin MSComctlLib.ListView lv_desgloce 
         Height          =   3045
         Left            =   60
         TabIndex        =   6
         Top             =   225
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   5371
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Orden"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Surtido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Empacado"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Negado"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Agente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cliente"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Existencias "
      Height          =   750
      Left            =   60
      TabIndex        =   7
      Top             =   1545
      Width           =   11505
      Begin VB.TextBox txt_disponible 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   8145
         TabIndex        =   5
         Top             =   285
         Width           =   1170
      End
      Begin VB.TextBox txt_apartado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   4590
         TabIndex        =   4
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txt_fisico 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   945
         TabIndex        =   3
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Disponibles:"
         Height          =   195
         Left            =   7260
         TabIndex        =   10
         Top             =   345
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Apartados:"
         Height          =   195
         Left            =   3810
         TabIndex        =   9
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fisíco:"
         Height          =   195
         Left            =   450
         TabIndex        =   8
         Top             =   345
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   795
      Width           =   11475
      Begin VB.TextBox txt_codigo 
         Height          =   330
         Left            =   105
         TabIndex        =   14
         Top             =   270
         Width           =   1875
      End
      Begin VB.TextBox txt_descripcion 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2025
         TabIndex        =   13
         Top             =   270
         Width           =   9300
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total en ordenes de surtido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6735
      TabIndex        =   15
      Top             =   6645
      Width           =   2415
   End
End
Attribute VB_Name = "frmexisten_rapidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
            sDsnName = "DSN=sqlsistema"
            sDriver = "SQL Server"
            dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)
            sDsnName = "sqlsistema"
            sDescription = "sqlsistema"
            sDriver = "SQL Server"
            sAttributes = "DSN=" & sDsnName & Chr(0)
            sAttributes = sAttributes & "Server=" + var_sr_reportes & Chr$(0)
            sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
            sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
            strAttributes = strAttributes & "UID=sa" & Chr$(0)
            strAttributes = strAttributes & "PWD=elia" & Chr$(0)
            dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)

            Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_existencias_rapidas.rpt")
            reporte.RecordSelectionFormula = "{VW_ORDEN_SURTIDO_EXISTENCIAS_RAPIDAS.VCHA_ART_ARTICULO_ID}= '" + Me.txt_codigo + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_existencias_rapidas_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 And Shift = 1 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_ubicaciones_almacen_detalle.rpt")
      reporte.RecordSelectionFormula = "{VW_UBICACIONES_ALMACEN_DETALLE.VCHA_ALM_ALMACEN_ID} = '" + Me.txt_clave_almacen + "' and {VW_UBICACIONES_ALMACEN_DETALLE.VCHA_aRT_ARTICULO_ID} = '" + Me.txt_codigo + "'"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de ubicaciones"
      frmvistasprevias.Show 1
      Set reporte = Nothing

   
      x = 1
      If x = 0 Then
      Dim list_item As ListItem
      frmubicaciones_almacen_detalle.txt_almacen = Me.txt_clave_almacen
      frmubicaciones_almacen_detalle.txt_nombre_almacen = Me.txt_nombre_almacen
      frmubicaciones_almacen_detalle.txt_articulo = Me.txt_codigo
      frmubicaciones_almacen_detalle.txt_nombre_articulo = Me.txt_descripcion
      rs.Open "select * from tb_ubicaciones_almacen_detalle where vcha_Art_articulo_id = '" + Me.txt_codigo + "' and vcha_Alm_almacen_id = '" + Me.txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = frmubicaciones_almacen_detalle.lv_ubicaciones.ListItems.Add(, , rs!VCHA_UBI_UBICACION)
               rs.MoveNext
         Wend
      End If
      rs.Close
      frmubicaciones_almacen_detalle.Show
      End If
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Me.frm_ubicaciones.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existen_rapidas)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub lv_ubicaciones_LostFocus()
   Me.frm_ubicaciones.Visible = False
End Sub

Private Sub txt_cantidad_ordenes_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Len(Trim(txt_clave_almacen)) > 0 Then
         rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
            txt_codigo.Enabled = True
            txt_codigo.SetFocus
            txt_clave_almacen.Enabled = False
            rs.Close
         Else
            rs.Close
            MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.lv_desgloce.ListItems.Clear
   Me.txt_fisico = ""
   Me.txt_apartado = ""
   Me.txt_disponible = ""
   Me.txt_cantidad_ordenes = ""
   Me.txt_descripcion = ""
   Me.txt_ubicacion_1 = ""
   Me.txt_ubicacion_2 = ""
   Me.txt_ubicacion_3 = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Len(Trim(txt_codigo)) > 0 Then
         Me.txt_cantidad_ordenes = "0"
         rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rsaux2.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               txt_codigo = rsaux2!VCHA_ART_ARTICULO_ID
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            End If
            rsaux2.Close
         End If
         rs.Close
         rs.Open "select * from tb_articulos where vcha_art_articulo_id ='" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         Dim list_item As ListItem
         If Not rs.EOF Then
            txt_descripcion = rs!vcha_Art_nombre_español
            rs.Close
            rsaux.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen + "' AND VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_cantidad_ordenes = "0"
               txt_fisico = Format(rsaux!floa_Exi_Cantidad, "###,###,##0.000")
               txt_apartado = Format(rsaux!FLOA_EXI_CANTIDAD_APARTADA, "###,###,##0.000")
               txt_disponible = Format(rsaux!floa_Exi_Cantidad_disponible, "###,###,##0.000")
               lv_desgloce.ListItems.Clear
               cnn.CommandTimeout = 360
               'rsaux2.Open "select inte_ors_orden_surtido,dtim_ors_fecha_carga,dtim_ors_fecha_caduca,floa_ors_cantidad_surtir,floa_ors_cantidad_surtida+isnull(floa_ors_cantidad_empacada,0)+isnull(floa_ors_cantidad_negada,0)as FLOA_ORS_CANTIDAD_SURTIDA,vcha_age_nombre,vcha_cli_nombre from vw_orden_surtido where floa_ors_cantidad_surtir > (floa_ors_cantidad_surtida + isnull(floa_ors_cantidad_empacada,0)+ isnull(floa_ors_cantidad_negada,0)) and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               'METODO NOT TAN NUEVO
               If UCase(Me.txt_clave_almacen) = "8" Or UCase(Me.txt_clave_almacen) = "PTTEX" Or UCase(Me.txt_clave_almacen) = "PTVH" Then
                  var_cadena = "SELECT dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA,dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CADUCA, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR,  dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_empacada,dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA,ISNULL(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA, 0) AS FLOA_ORS_CANTIDAD_negada,  dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID FROM         dbo.TB_AGENTES INNER JOIN "
                  var_cadena = var_cadena + " dbo.TB_CLIENTES ON dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID = dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_DET_ORDEN_SURTIDO INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON  dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO AND  dbo.TB_DET_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_EMP_EMPRESA_ID AND  dbo.TB_DET_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_UOR_UNIDAD_ID AND  dbo.TB_DET_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_ALM_ALMACEN_ID ON  dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR > isnull(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_negada,0) + isnull(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SALIDA,0)) AND (dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "')"
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  
                  While Not rsaux2.EOF
                        Set list_item = lv_desgloce.ListItems.Add(, , rsaux2!INTE_ORS_ORDEN_SURTIDO)
                        list_item.SubItems(1) = Format(IIf(IsNull(rsaux2!DTIM_ORS_FECHA_CARGA), "", rsaux2!DTIM_ORS_FECHA_CARGA), "Short Date")
                        'list_item.SubItems(2) = IIf(IsNull(rsaux2!DTIM_ORS_FECHA_CADUCA), "", rsaux2!DTIM_ORS_FECHA_CADUCA)
                        list_item.SubItems(2) = IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIR), Format(0, "###,###,##0.000"), Format(rsaux2!FLOA_ORS_CANTIDAD_SURTIR, "###,###,##0.000"))
                        list_item.SubItems(3) = IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIDA), Format(0, "###,###,##0.000"), Format(rsaux2!FLOA_ORS_CANTIDAD_SURTIDA, "###,###,##0.000"))
                        list_item.SubItems(4) = IIf(IsNull(rsaux2!floa_ors_Cantidad_empacada), Format(0, "###,###,##0.000"), Format(rsaux2!floa_ors_Cantidad_empacada, "###,###,##0.000"))
                        list_item.SubItems(5) = IIf(IsNull(rsaux2!floa_ors_cantidad_negada), Format(0, "###,###,##0.000"), Format(rsaux2!floa_ors_cantidad_negada, "###,###,##0.000"))
                        list_item.SubItems(6) = IIf(IsNull(rsaux2!VCHA_AGE_NOMBRE), "", rsaux2!VCHA_AGE_NOMBRE)
                        list_item.SubItems(7) = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                        list_item.SubItems(8) = ""
                        Me.txt_cantidad_ordenes = CDbl(Me.txt_cantidad_ordenes) + IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux2!FLOA_ORS_CANTIDAD_SURTIR)
                        rsaux2.MoveNext
                  Wend
                  rsaux2.Close
               
               
                     
                  'var_cadena = " SELECT     TOP 100 PERCENT dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA, dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CADUCA, "
                  'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA,                      dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_EMPACADA, dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA_CERRADO FROM dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON"
                  'var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON  "
                  'var_cadena = var_cadena + " dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID whERE     (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR = dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA + ISNULL(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_NEGADA,0)) AND (dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS = 'S') AND (dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR > 0) AND (dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "') AND (dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA >= CONVERT(DATETIME, '2008-08-15', 102)) ORDER BY dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO"
                  'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  '
                  '
                  'While Not rsaux2.EOF
                  '      Set list_item = lv_desgloce.ListItems.Add(, , rsaux2!INTE_ORS_ORDEN_SURTIDO)
                  '      list_item.SubItems(1) = IIf(IsNull(rsaux2!DTIM_ORS_FECHA_CARGA), "", rsaux2!DTIM_ORS_FECHA_CARGA)
                  '      'list_item.SubItems(2) = IIf(IsNull(rsaux2!DTIM_ORS_FECHA_CADUCA), "", rsaux2!DTIM_ORS_FECHA_CADUCA)
                  '      list_item.SubItems(2) = IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIR), Format(0, "###,###,##0.000"), Format(rsaux2!FLOA_ORS_CANTIDAD_SURTIR, "###,###,##0.000"))
                  '      list_item.SubItems(3) = IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIDA), Format(0, "###,###,##0.000"), Format(rsaux2!FLOA_ORS_CANTIDAD_SURTIDA, "###,###,##0.000"))
                  '      list_item.SubItems(4) = IIf(IsNull(rsaux2!floa_ors_Cantidad_empacada), Format(0, "###,###,##0.000"), Format(rsaux2!floa_ors_Cantidad_empacada, "###,###,##0.000"))
                  '      list_item.SubItems(5) = IIf(IsNull(rsaux2!floa_ors_cantidad_negada), Format(0, "###,###,##0.000"), Format(rsaux2!floa_ors_cantidad_negada, "###,###,##0.000"))
                  '      list_item.SubItems(6) = IIf(IsNull(rsaux2!vcha_age_nombre), "", rsaux2!vcha_age_nombre)
                  '      list_item.SubItems(7) = IIf(IsNull(rsaux2!vcha_cli_nombre), "", rsaux2!vcha_cli_nombre)
                  '      list_item.SubItems(8) = "*"
                  '      Me.txt_cantidad_ordenes = CDbl(Me.txt_cantidad_ordenes) + IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux2!FLOA_ORS_CANTIDAD_SURTIR)
                  '      rsaux2.MoveNext
                  'Wend
                  'rsaux2.Close
                  'Me.txt_cantidad_ordenes = Format(CDbl(Me.txt_cantidad_ordenes), "###,###,##0.000")
                  'For var_j = 1 To Me.lv_desgloce.ListItems.Count
                  '    lv_desgloce.ListItems.Item(var_j).Selected = True
                  '    If lv_desgloce.selectedItem.SubItems(8) = "*" Then
                  '       lv_desgloce.ListItems.Item(var_j).Selected = True
                  '       lv_desgloce.ListItems.Item(var_j).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(1).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(2).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(3).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(4).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(5).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(6).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(7).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(8).Bold = True
                  '       'lv_desgloce.ListItems.Item(var_j).ListSubItems(9).Bold = True
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HFF0000
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HFF0000
                  '       'lv_desgloce.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HFF0000
                  '    Else
                  '       lv_desgloce.ListItems.Item(var_j).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(1).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(2).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(3).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(4).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(5).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(6).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(7).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(8).Bold = False
                  '       'lv_desgloce.ListItems.Item(var_j).ListSubItems(9).Bold = False
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(7).ForeColor = &H80000012
                  '       lv_desgloce.ListItems.Item(var_j).ListSubItems(8).ForeColor = &H80000012
                  '       'lv_desgloce.ListItems.Item(var_j).ListSubItems(9).ForeColor = &H80000012
                  '    End If
                  'Next var_j
               End If
            Else
               lv_desgloce.ListItems.Clear
               txt_fisico = 0
               txt_apartado = 0
               txt_disponible = 0
            End If
            rsaux.Close
            rsaux.Open "select * from tb_ubicaciones_almacen where vcha_alm_almacen_id = '" + Me.txt_clave_almacen + "' and vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_ubicacion_1 = IIf(IsNull(rsaux!vcha_ubi_ubicacion_1), "", rsaux!vcha_ubi_ubicacion_1)
               Me.txt_ubicacion_2 = IIf(IsNull(rsaux!vcha_ubi_ubicacion_2), "", rsaux!vcha_ubi_ubicacion_2)
               Me.txt_ubicacion_3 = IIf(IsNull(rsaux!vcha_ubi_ubicacion_3), "", rsaux!vcha_ubi_ubicacion_3)
            Else
               Me.txt_ubicacion_1 = ""
               Me.txt_ubicacion_2 = ""
               Me.txt_ubicacion_3 = ""
            End If
            rsaux.Close
         Else
            rs.Close
            'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
      End If
      
   End If
End Sub

Private Sub txt_ubicacion_1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_ubicaciones.Visible = True
      Me.lv_ubicaciones.SetFocus
   End If
End Sub
