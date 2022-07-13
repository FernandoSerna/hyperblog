VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsalidas_numero_serie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salidas de Numero de serie"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1290
      TabIndex        =   0
      Top             =   690
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   1
         Top             =   480
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
         TabIndex        =   2
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   720
      TabIndex        =   13
      Top             =   885
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   14
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7170
      Picture         =   "frmsalidas_numero_serie.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmsalidas_numero_serie.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmsalidas_numero_serie.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmsalidas_numero_serie.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Buscar Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmsalidas_numero_serie.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   570
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   0
      Left            =   5760
      TabIndex        =   4
      Top             =   1110
      Width           =   1770
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2820
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   75
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_numero_serie.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   16
      Top             =   975
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      Height          =   1245
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   1110
      Width           =   5610
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   825
         TabIndex        =   32
         Top             =   810
         Width           =   1125
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1980
         TabIndex        =   20
         Top             =   825
         Width           =   3570
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   450
         Width           =   3570
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   825
         TabIndex        =   18
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   23
         Top             =   885
         Width           =   525
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   510
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4965
      Left            =   120
      TabIndex        =   24
      Top             =   2295
      Width           =   7425
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2085
         TabIndex        =   26
         Top             =   495
         Width           =   2640
      End
      Begin VB.TextBox txt_placa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5655
         TabIndex        =   25
         Top             =   555
         Width           =   1515
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3780
         Left            =   45
         TabIndex        =   27
         Top             =   1110
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   6668
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
            Text            =   "Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número de serie"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   1940
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código o Número de Serie:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   675
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Left            =   4950
         TabIndex        =   28
         Top             =   675
         Width           =   450
      End
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      TabIndex        =   31
      Top             =   75
      Width           =   7335
   End
End
Attribute VB_Name = "frmsalidas_numero_serie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen_Destino As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_numero_causa As Double
Dim var_elimina As Boolean
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_año As Integer
Dim var_suma_cantidad As Double
Dim var_cantidad_llegar As Double
Dim var_cantidad As Double
Dim var_renglon As Double
Dim var_codigo As String
Dim var_empresa_salida As String
Dim var_unidad_salida As String
Dim var_almacen_salida As String
Dim var_movimiento_salida As String
Dim var_numero_salida As Double
Dim var_tipo_lista As Integer
Dim var_codigo_serie_existe As Boolean



Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.Item(var_i).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_entradas.ListItems.Item(var_i).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_entradas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.Item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   lv_entradas.Refresh
End Sub




Private Sub cmd_buscar_Click()
   var_ventana = 1
   frm_busqueda.Visible = True
   txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_posible_cerrar_movimiento As Integer
   var_posible_cerrar_movimiento = 1
   
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_movimientos & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_salida_numero_serie.rpt")
         reporte.RecordSelectionFormula = "{VW_SALIDAS_NUMERO_SERIE.vcha_EXI_EMPRESA_SALIDA} = '" + var_empresa + "' and {VW_SALIDAS_NUMERO_SERIE.VCHA_EXI_ALMACEN_SALIDA} = '" + var_almacen_Destino + "' AND {VW_SALIDAS_NUMERO_SERIE.VCHA_EXI_MOVIMIENTO_SALIDA} = '" + var_clave_movimiento + "' AND {VW_SALIDAS_NUMERO_SERIE.INTE_EXI_NUMERO_SALIDA} = " + Str(var_numero_folio) + " AND {VW_SALIDAS_NUMERO_SERIE.VCHA_EXI_UNIDAD_SALIDA} = '" + var_unidad_organizacional + "'"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      Else
         var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            cnn.BeginTrans
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from tb_existencias_series where vcha_exi_empresa_salida = '" + var_empresa + "' and inte_exi_numero_salida = " + CStr(var_numero_folio) + " and vcha_exi_movimiento_salida = '" + var_clave_movimiento + "' and vcha_exi_empresa_salida = '" + var_empresa + "' and vcha_exi_unidad_salida = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_cantidad = IIf(IsNull(rs!floa_Exi_Cantidad), "", rs!floa_Exi_Cantidad)
                  rsaux5.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     var_costo = IIf(IsNull(rsaux5!mone_Art_costo_estandar), 0, rsaux5!mone_Art_costo_estandar)
                     var_precio = IIf(IsNull(rsaux5!mone_art_precio_base), 0, rsaux5!mone_art_precio_base)
                  Else
                     var_costo = 0
                     var_precio = 0
                  End If
                  rsaux5.Close
                  rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + rs!vcha_exi_empresa_salida + "', '" + rs!vcha_exi_unidad_salida + "', '" + rs!vcha_exi_almacen_salida + "', '" + rs!vcha_exi_movimiento_salida + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(var_precio) + ", 2005)", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
            var_estatus_movimiento = "I"
            cnn.CommitTrans
            Set reporte = appl.OpenReport(App.Path + "\rep_salida_numero_serie.rpt")
            reporte.RecordSelectionFormula = "{VW_SALIDAS_NUMERO_SERIE.vcha_EXI_EMPRESA_SALIDA} = '" + var_empresa + "' and {VW_SALIDAS_NUMERO_SERIE.VCHA_EXI_ALMACEN_SALIDA} = '" + var_almacen_Destino + "' AND {VW_SALIDAS_NUMERO_SERIE.VCHA_EXI_MOVIMIENTO_SALIDA} = '" + var_clave_movimiento + "' AND {VW_SALIDAS_NUMERO_SERIE.INTE_EXI_NUMERO_SALIDA} = " + Str(var_numero_folio) + " AND {VW_SALIDAS_NUMERO_SERIE.VCHA_EXI_UNIDAD_SALIDA} = '" + var_unidad_organizacional + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            txt_codigo.Enabled = False
            txt_foco.Enabled = False
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   txt_almacen = ""
   txt_nombre_almacen = ""
   var_ventana = 0
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   lv_entradas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_referencia = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   txt_almacen.Enabled = True
   Me.txt_cliente.Enabled = False
   txt_almacen.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show 1
   End If
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 66 Then
      cmd_buscar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 67 Then
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And var_ventana = 0 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 0
   Left = 1500
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
   var_clave_moneda = ""
   If Not rs.EOF Then
      var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   End If
   rs.Close
   var_ventana = 0
   var_estatus_movimiento = ""
   frm_busqueda.Visible = False
   Me.txt_placa.Enabled = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_salidas_sin_comparacion)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         var_si = MsgBox("¿Desea eliminar el artículo del movimiento?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            If Trim(lv_entradas.selectedItem.SubItems(1)) <> "" Then
               rs.Open "select * from TB_EXISTENCIAS_SERIES where vcha_Art_numero_Serie = '" + lv_entradas.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
               var_codigo = ""
               var_empresa_salida = ""
               var_unidad_salida = ""
               var_almacen_salida = ""
               var_movimiento_salida = ""
               var_numero_salida = 0
               var_codigo = IIf(IsNull(rs!vcha_Art_articulo_id), "", rs!vcha_Art_articulo_id)
               var_empresa_salida = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
               var_unidad_salida = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
               var_almacen_salida = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
               var_movimiento_salida = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
               var_numero_salida = IIf(IsNull(rs!INTE_EMO_NUMERO), "", rs!INTE_EMO_NUMERO)
               var_numero_serie = IIf(IsNull(rs!vcha_Art_numero_Serie), "", rs!vcha_Art_numero_Serie)
               rs.Close
               rsaux.Open "UPDATE TB_EXISTENCIAS_SERIES SET VCHA_EXI_EMPRESA_SALIDA = '', VCHA_EXI_UNIDAD_SALIDA = '', VCHA_EXI_ALMACEN_SALIDA = '', VCHA_EXI_MOVIMIENTO_SALIDA = '', INTE_EXI_NUMERO_SALIDA = 0, floa_exi_Cantidad = 0, vcha_exi_placa = '' WHERE vcha_emp_empresa_id = '" + var_empresa_salida + "' and vcha_uor_unidad_id = '" + var_unidad_salida + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_salida + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_salida) + " AND VCHA_aRT_ARTICULO_ID = '" + var_codigo + "' AND VCHA_ART_NUMERO_SERIE = '" + Me.lv_entradas.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
               lv_entradas.ListItems.Remove (lv_entradas.selectedItem.Index)
            Else
               If CDbl(lv_entradas.selectedItem.SubItems(3)) = 0 Then
                  MsgBox "No se puede eliminar", vbOKOnly, "ATENCION"
               Else
                  rsaux.Open "UPDATE TB_EXISTENCIAS_SERIES SET floa_exi_Cantidad =  isnull(floa_exi_Cantidad,0) - 1 where VCHA_EXI_EMPRESA_SALIDA = '" + var_empresa + "' AND VCHA_EXI_UNIDAD_SALIDA = '" + var_unidad_organizacional + "' AND VCHA_EXI_ALMACEN_SALIDA = '" + var_almacen_Destino + "' AND VCHA_EXI_MOVIMIENTO_SALIDA = '" + var_clave_movimiento + "' AND INTE_EXI_NUMERO_SALIDA = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  lv_entradas.selectedItem.SubItems(3) = CDbl(lv_entradas.selectedItem.SubItems(3)) - 1
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_almacen = lv_lista.selectedItem
            txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
         Else
            txt_almacen = ""
            txt_nombre_almacen = ""
         End If
         txt_almacen.SetFocus
         frm_lista.Visible = False
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            Me.txt_cliente = lv_lista.selectedItem
            Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            Me.txt_cliente = ""
            Me.txt_nombre_cliente = ""
         End If
         Me.txt_cliente.SetFocus
         var_tipo_lista = 0
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      'rs.Open "select distinct vcha_cli_nombre from vw_establecimientos where vcha_esb_establecimiento_id = '" + txt_establecimiento + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes"
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

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen) <> "" Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "'  AND VCHA_ALM_ALMACEN_ID = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         txt_almacen.Enabled = False
         txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
         var_almacen_Destino = txt_almacen
         Me.txt_cliente.Enabled = True
         
      Else
         MsgBox "Clave de almacen Incorrecta", vbOKOnly, "ATENCION"
         txt_almacen = ""
         txt_nombre_almacen = ""
         Me.txt_cliente.Enabled = False
      End If
      If rs.State = 1 Then
         rs.Close
      End If
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
            If var_movimiento_bloqueado = 0 Then
               var_almacen_destino_tem = rs!VCHA_ALM_ALMACEN_ID
               var_posible = 1
               If var_tipo_permiso = 1 Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               If var_posible = 1 Then
                  var_estatus_movimiento = rs!char_Emo_estatus
                  var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
                  txt_almacen = rs!VCHA_ALM_ALMACEN_ID
                  txt_almacen.Enabled = False
                  txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                  rsaux5.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     Me.txt_nombre_cliente = IIf(IsNull(rsaux5!VCHA_CLI_NOMBRE), "", rsaux5!VCHA_CLI_NOMBRE)
                  End If
                  rsaux5.Close
                  txt_cliente.Enabled = False
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(3).Value
                  txt_nombre_almacen.Text = rsaux(3).Value
                  rsaux.Close
                  rsaux.Open "select * from tb_existencias_series where vcha_exi_empresa_salida = '" + var_empresa + "' and inte_exi_numero_salida = " + txt_busqueda_folio + " and vcha_exi_movimiento_salida = '" + var_clave_movimiento + "' and vcha_exi_empresa_salida = '" + var_empresa + "' and vcha_exi_unidad_salida = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_entradas.ListItems.Add(, , IIf(IsNull(rsaux!vcha_Art_articulo_id), "", rsaux!vcha_Art_articulo_id))
                           list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Art_numero_Serie), "", rsaux!vcha_Art_numero_Serie)
                           list_item.SubItems(2) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(3) = IIf(IsNull(rsaux!floa_Exi_Cantidad), "", rsaux!floa_Exi_Cantidad)
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                  End If
                  rsaux.Close
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     txt_codigo.Enabled = False
                     txt_foco.Enabled = False
                     txt_placa.Enabled = False
                  Else
                     txt_foco.Enabled = False
                     txt_codigo.Enabled = True
                     txt_placa.Enabled = False
                  End If
               Else
                  MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento esta siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      var_ventana = 0
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_eliminar) Then
         Dim var_posible_eliminar As Boolean
         var_cantidad_eliminar = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If var_posible_eliminar = True Then
            var_inserta = False
            rsaux.Open "UPDATE TB_TEMPORAL_SALIDAS SET FLOA_SAL_CANTIDAD = ISNULL(FLOA_SAL_CANTIDAD,0) - " + txt_cantidad_eliminar + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio) + " AND VCHA_ART_ARTICULO_ID= '" + lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            'var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, lv_entradas.SelectedItem, 0 - Val(txt_cantidad_eliminar))
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
            var_renglon = lv_entradas.selectedItem.Index
            Call ilumina_grid
         Else
            MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devolución seleccionada", vbOKOnly, "ATENCION"
         End If
         var_ventana = 0
         frm_eliminar.Visible = False
         txt_codigo.SetFocus
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = 1#
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_cantidad) <> "" Then
         var_cantidad_leida = txt_cantidad
         txt_foco.Enabled = True
         txt_foco.SetFocus
         lbl_cantidad.Visible = False
         txt_cantidad.Visible = False
      End If
   End If
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and char_tpe_tipo_pedido_id = 'T' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 2
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

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_cliente) <> "" Then
         rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            Me.txt_codigo.Enabled = True
            Me.txt_codigo.SetFocus
         Else
            Me.txt_nombre_cliente = ""
         End If
         rs.Close
      End If
   End If
End Sub


Private Sub txt_cliente_LostFocus()
   If var_tipo_lista <> 2 Then
      Me.txt_cliente.Enabled = False
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   Dim var_requiere_numero_Serie As Integer
   txt_codigo = Trim(txt_codigo)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from TB_EXISTENCIAS_SERIES where vcha_Art_numero_Serie = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      var_codigo = ""
      var_empresa_salida = ""
      var_unidad_salida = ""
      var_almacen_salida = ""
      var_movimiento_salida = ""
      var_numero_salida = 0
      If Not rs.EOF Then
         var_codigo_serie_existe = True
         If Trim(IIf(IsNull(rs!vcha_exi_empresa_salida), "", rs!vcha_exi_empresa_salida)) = "" Then
            var_codigo = IIf(IsNull(rs!vcha_Art_articulo_id), "", rs!vcha_Art_articulo_id)
            var_empresa_salida = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
            var_unidad_salida = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
            var_almacen_salida = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
            var_movimiento_salida = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
            var_numero_salida = IIf(IsNull(rs!INTE_EMO_NUMERO), "", rs!INTE_EMO_NUMERO)
            Me.txt_placa.Enabled = True
            Me.txt_placa.SetFocus
         Else
            MsgBox "El número de serie ya fue cargado", vbOKOnly, "ATENCION"
         End If
      Else
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_requiere_numero_Serie = IIf(IsNull(rsaux!inte_art_numero_serie), 0, rsaux!inte_art_numero_serie)
            If var_requiere_numero_Serie > 0 Then
               MsgBox "El artículo requiere número de serie", vbOKOnly, "ATENCION"
            Else
               var_codigo = IIf(IsNull(rsaux!vcha_Art_articulo_id), "", rsaux!vcha_Art_articulo_id)
               var_codigo_serie_existe = False
               Me.txt_placa.Enabled = False
               Me.txt_foco.Enabled = True
               Me.txt_foco.SetFocus
            End If
         Else
            MsgBox "El número de serie no se encuentra cargado en el sistema", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
      rs.Close
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      bandera_suma = False
      If var_primera_vez = True Then
         var_inserta = False
         var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, txt_cliente, "", var_almacen_Destino, "", "", var_clave_usuario_global, fun_NombrePc, 0, "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
         var_numero_folio = var_numero_folio_regreso
         txt_folio = var_numero_folio
         var_primera_vez = False
      End If
      Dim var_si_encontro As Integer
      var_si_encontro = 0
      If var_codigo_serie_existe = False Then
         rsaux.Open "select * from  TB_EXISTENCIAS_SERIES where VCHA_EXI_EMPRESA_SALIDA = '" + var_empresa + "' AND VCHA_EXI_UNIDAD_SALIDA = '" + var_unidad_organizacional + "' AND VCHA_EXI_ALMACEN_SALIDA = '" + var_almacen_Destino + "' AND VCHA_EXI_MOVIMIENTO_SALIDA = '" + var_clave_movimiento + "' AND INTE_EXI_NUMERO_SALIDA = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux2.Open "update  TB_EXISTENCIAS_SERIES  set FLOA_eXI_CANTIDAD = ISNULL(FLOA_EXI_CANTIDAD,0) + 1 where VCHA_EXI_EMPRESA_SALIDA = '" + var_empresa + "' AND VCHA_EXI_UNIDAD_SALIDA = '" + var_unidad_organizacional + "' AND VCHA_EXI_ALMACEN_SALIDA = '" + var_almacen_Destino + "' AND VCHA_EXI_MOVIMIENTO_SALIDA = '" + var_clave_movimiento + "' AND INTE_EXI_NUMERO_SALIDA = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            var_si_encontro = 1
         Else
            rsaux2.Open "INSERT INTO TB_EXISTENCIAS_SERIES (VCHA_EXI_EMPRESA_SALIDA, VCHA_EXI_UNIDAD_SALIDA, VCHA_EXI_ALMACEN_SALIDA, VCHA_EXI_MOVIMIENTO_SALIDA, INTE_EXI_NUMERO_SALIDA, vcha_Art_articulo_id, floa_exi_cantidad, vcha_exi_placa)  VALUES  ('" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_Destino + "','" + var_clave_movimiento + "'," + CStr(var_numero_folio) + ",'" + Me.txt_codigo + "',1,'')", cnn, adOpenDynamic, adLockOptimistic
            var_si_encontro = 0
         End If
         rsaux.Close
      Else
         var_si_encontro = 0
         rsaux.Open "UPDATE TB_EXISTENCIAS_SERIES SET VCHA_EXI_EMPRESA_SALIDA = '" + var_empresa + "', VCHA_EXI_UNIDAD_SALIDA = '" + var_unidad_organizacional + "', VCHA_EXI_ALMACEN_SALIDA = '" + var_almacen_Destino + "', VCHA_EXI_MOVIMIENTO_SALIDA = '" + var_clave_movimiento + "', INTE_EXI_NUMERO_SALIDA = " + CStr(var_numero_folio) + ", floa_exi_cantidad = 1, vcha_exi_placa = '" + Me.txt_placa + "' WHERE vcha_emp_empresa_id = '" + var_empresa_salida + "' and vcha_uor_unidad_id = '" + var_unidad_salida + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_salida + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_salida) + " AND VCHA_aRT_ARTICULO_ID = '" + var_codigo + "' AND VCHA_ART_NUMERO_SERIE = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      var_descripcion_articulo = rs!vcha_art_nombre_español
      If var_codigo_serie_existe = False Then
          If var_si_encontro = 0 Then
             Set list_item = lv_entradas.ListItems.Add(, , Trim(var_codigo))
             list_item.SubItems(2) = var_descripcion_articulo
             list_item.SubItems(3) = 1
          Else
            Set itmfound = lv_entradas.findItem(var_codigo, lvwText, , lvwPartial)
            itmfound.EnsureVisible
            itmfound.Selected = True
            lv_entradas.selectedItem.SubItems(3) = CDbl(lv_entradas.selectedItem.SubItems(3)) + 1
            lv_entradas.selectedItem.SubItems(2) = var_descripcion_articulo
          End If
      Else
         Set list_item = lv_entradas.ListItems.Add(, , Trim(var_codigo))
         list_item.SubItems(1) = txt_codigo
         list_item.SubItems(2) = var_descripcion_articulo
         list_item.SubItems(3) = 1
      End If
      var_renglon = lv_entradas.ListItems.Count
      Call ilumina_grid
      txt_codigo.SetFocus
      txt_placa = ""
   End If
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_almacen.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Else
            rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Almacenes"
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
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_almacen) <> "" Then
         If txt_cliente.Enabled = True Then
            txt_cliente.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Len(Trim(txt_referencia)) > 0 Then
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
         txt_referencia.Enabled = False
      Else
         MsgBox "Debe introducir una referencia", vbOKOnly, "ATENCION"
      End If
   End If
End Sub


Private Sub txt_placa_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_foco.Enabled = True
      Me.txt_foco.SetFocus
   End If
End Sub
