VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsalidas_transformacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida de transformación"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   10305
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3465
      Width           =   1125
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2250
      TabIndex        =   31
      Top             =   150
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   540
      TabIndex        =   28
      Top             =   780
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   29
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
         TabIndex        =   30
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmsalidas_transformacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   690
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   2
      Left            =   90
      TabIndex        =   26
      Top             =   525
      Width           =   9900
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9600
      Picture         =   "frmsalidas_transformacion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   690
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmsalidas_transformacion.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   690
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmsalidas_transformacion.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   690
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   90
      TabIndex        =   25
      Top             =   945
      Width           =   9900
   End
   Begin VB.Frame Frame3 
      Caption         =   " Movimiento "
      Height          =   780
      Left            =   120
      TabIndex        =   22
      Top             =   1095
      Width           =   9795
      Begin VB.CheckBox chk_uno_a_uno 
         Caption         =   "Uno a uno"
         Height          =   345
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
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
         Left            =   7845
         TabIndex        =   7
         Top             =   195
         Width           =   1890
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   345
         Left            =   1755
         TabIndex        =   5
         Top             =   240
         Width           =   3765
      End
      Begin VB.TextBox txt_almacen 
         Height          =   345
         Left            =   855
         TabIndex        =   4
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   7335
         TabIndex        =   24
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacén:"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   315
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo general "
      Height          =   990
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1890
      Width           =   9795
      Begin VB.TextBox txt_cantidad_general 
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
         Left            =   7455
         TabIndex        =   10
         Top             =   300
         Width           =   2250
      End
      Begin VB.TextBox txt_nombre_articulo_general 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2475
         TabIndex        =   9
         Top             =   300
         Width           =   4965
      End
      Begin VB.TextBox txt_codigo_general 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   90
         TabIndex        =   8
         Top             =   300
         Width           =   2370
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4365
      Left            =   120
      TabIndex        =   14
      Top             =   2865
      Width           =   9795
      Begin VB.TextBox txt_cantidad 
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
         Left            =   5865
         TabIndex        =   12
         Top             =   495
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   6285
         TabIndex        =   15
         Top             =   2385
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   16
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   17
            Top             =   15
            Width           =   2895
         End
      End
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
         Left            =   1560
         TabIndex        =   11
         Top             =   465
         Width           =   3390
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   3270
         Left            =   15
         TabIndex        =   13
         Top             =   1035
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   5768
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
            Text            =   "   Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   9349
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   9735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5115
         TabIndex        =   18
         Top             =   615
         Width           =   675
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
      Left            =   75
      TabIndex        =   27
      Top             =   0
      Width           =   9795
   End
End
Attribute VB_Name = "frmsalidas_transformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_cantidad_multibondeados As Double
Dim var_kanban As String
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
Dim var_numero_causa As Integer
Dim var_elimina As Boolean
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_año As Integer
Dim var_suma_cantidad As Double
Dim var_cantidad_llegar As Double
Dim var_cantidad As Double
Dim var_renglon As Double
Dim var_tipo_lista As Integer
Dim var_cadena_conexion As String
Dim cnn_traspaso_intecomañia As ADODB.Connection




Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda_folio = ""
   Me.txt_busqueda_folio.SetFocus
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
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   var_almacen_Destino = txt_almacen
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS_transformacion.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Str(Me.txt_folio) + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = '" + Me.txt_almacen + "' and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '" + var_empresa + "'"
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
            var_numero_folio = CDbl(Me.txt_folio)
            rs.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "INSERT INTO TB_SALIDAS (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_SAL_numero, vcha_art_articulo_id, floa_SAL_cantidad, floa_SAL_costo, floa_SAL_precio, INTE_SAL_AÑO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_almacen + "', '" + var_clave_movimiento + "', " + CStr(Me.txt_folio) + ", '" + Me.txt_codigo_general + "', " + CStr(CDbl(Me.txt_cantidad_general)) + ", " + CStr(IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)) + " , " + CStr(IIf(IsNull(rs!mone_Art_precio_base), 0, rs!mone_Art_precio_base)) + ", 2005)", cnn, adOpenDynamic, adLockOptimistic
            rs.Close
            Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + txt_almacen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_inserta = False
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_ent_Cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            var_estatus_movimiento = "I"
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, txt_almacen, var_clave_movimiento, var_numero_folio, "", Now, 1)
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, txt_almacen, var_clave_movimiento, var_numero_folio, "I", Now, 1)
            cnn.CommitTrans
            Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS_transformacion.rpt")
            reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Me.txt_folio + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = '" + txt_almacen + "'  and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '" + var_empresa + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            txt_codigo.Enabled = False
            txt_foco.Enabled = False
            rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_empresa = "31" Then
      var_si = MsgBox("¿Desea validar los artículos nuevos?", vbYesNo, "ATENCION")
   Else
      var_si = 0
   End If
   If var_si = 6 Then
      rs.Open "SELECT ART_CODIGO AS VCHA_ART_ARTICULO_ID, ART_GTIN AS VCHA_EQU_CODIGO_EQUIVALENTE, EXPR1 AS VCHA_ART_NOMBRE_ESPAÑOL, ART_ULTIMOCOSTO AS MONE_ART_COSTO_ESTANDAR, LPA_PRECIOVENTA /1.16 AS MONE_ART_PRECIO_BASE FROM AVL_PRECIOS", cnn_compucaja, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            rsaux.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "D" Then
                  If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "P" Then
                     If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "R" Then
                        var_cadena = "insert into tb_Articulos (vcha_art_articulo_id, vcha_Art_nombre_Español, mone_art_costo_estandar, mone_art_precio_base, dtim_Art_fecha_alta, vcha_Art_catalogo_inicio, vcha_Art_catalogo_vigente, vcha_lic_licencia_id, vcha_Art_numero_lic, vcha_tal_talla_id, vcha_uni_unidad_id, inte_art_detenido, vcha_emp_empresa_id )"
                        'MsgBox rs!vcha_Art_articulo_id
                        var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "','" + rs!vcha_art_nombre_Español + "'," + CStr(rs!mone_Art_costo_estandar) + "," + CStr(rs!mone_Art_precio_base) + ",getdate(),'CANTIA','CANTIA','SIN LICENCIA','','UNI','01',0,'31')"
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
               End If
            End If
            rsaux.Close
            var_si_equivalencia = IIf(IsNull(rs!vcha_equ_codigo_equivalente), "", rs!vcha_equ_codigo_equivalente)
            If var_si_equivalencia <> "" Then
               rsaux.Open "SELECT * FROM tb_equivalencias WHERE vcha_equ_codigo_equivalente = '" + rs!vcha_equ_codigo_equivalente + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "D" Then
                     If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "P" Then
                        If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "R" Then
                           var_cadena = "insert into tb_equivalencias (vcha_art_articulo_id, vcha_equ_codigo_equivalente)"
                           var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "','" + rs!vcha_equ_codigo_equivalente + "')"
                           rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        End If
                     End If
                  End If
               End If
               rsaux.Close
            End If
            
            rsaux.Open "SELECT * FROM tb_Detalle_lista_precios WHERE VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "' and vcha_lis_lista_precios_id = '01'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "D" Then
                  If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "P" Then
                     If Mid(rs!vcha_Art_Articulo_id, 1, 1) <> "R" Then
                        var_cadena = "insert into tb_Detalle_lista_precios (vcha_art_articulo_id, vcha_lis_lista_precios_id, floa_dli_precio)"
                        'MsgBox rs!vcha_Art_articulo_id
                        var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "','01'," + CStr(rs!mone_Art_precio_base) + ") "
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
               End If
            End If
            rsaux.Close
            
            rs.MoveNext
      Wend
      rs.Close
      MsgBox "Se termino la importación de los artículos", vbOKOnly, "ATENCION"
   End If
   Me.txt_almacen = ""
   Me.txt_nombre_almacen = ""
   Me.txt_folio = ""
   Me.txt_codigo = ""
   Me.txt_cantidad = ""
   Me.lv_lista.ListItems.Clear
   Me.chk_uno_a_uno = 0
   Me.txt_codigo_general = ""
   Me.txt_nombre_articulo_general = ""
   Me.txt_cantidad_general = ""
   Me.txt_almacen.Enabled = True
   Me.txt_nombre_almacen.Enabled = True
   Me.lv_salidas.ListItems.Clear
   Me.txt_almacen.SetFocus
   var_estatus_movimiento = ""
   var_primera_vez = True
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 800
   Me.frm_lista.Visible = False
   Me.frm_eliminar.Visible = False
   Me.frm_busqueda.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_salidas_proveedor)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
    Me.txt_almacen = Me.lv_lista.selectedItem
    Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
    Me.txt_almacen.SetFocus
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub


Private Sub lv_salidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If Me.lv_salidas.ListItems.Count > 0 Then
      If KeyCode = 114 Then
         Me.frm_eliminar.Visible = True
         Me.txt_cantidad_eliminar = ""
         Me.txt_cantidad_eliminar.SetFocus
      End If
   End If
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
      If Not rs.EOF Then
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
      Else
         MsgBox "No existen almacenes para este movimiento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.txt_nombre_almacen.SetFocus
    End If
End Sub

Private Sub txt_almacen_LostFocus()
   If Me.txt_almacen = "" Then
      Me.txt_nombre_almacen = ""
   Else
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         Me.txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
         Me.txt_codigo_general.Enabled = True
         Me.txt_nombre_articulo_general.Enabled = True
         Me.txt_almacen.Enabled = False
         Me.txt_nombre_almacen.Enabled = False
      Else
         MsgBox "Clave de almacén incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
   
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda_folio) Then
         rs.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_EMP_eMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.lv_salidas.ListItems.Clear
            var_numero_folio = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
            Me.txt_folio = var_numero_folio
            Me.txt_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
            var_almacen_Destino = txt_almacen

            rsaux.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_nombre_almacen = IIf(IsNull(rsaux!VCHA_ALM_NOMBRE), "", rsaux!VCHA_ALM_NOMBRE)
               rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + IIf(IsNull(rs!vcha_Emo_referencia), "", rs!vcha_Emo_referencia) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  rsaux3.Open "SELECT * FROM TB_tEMPORAL_ENTRADAS WHERE  VCHA_EMP_eMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO = " + Me.txt_busqueda_folio, cnn, adOpenDynamic, adLockOptimistic
                  Me.txt_codigo_general = rsaux1!vcha_Art_Articulo_id
                  Me.txt_nombre_articulo_general = IIf(IsNull(rsaux1!vcha_art_nombre_Español), "", rsaux1!vcha_art_nombre_Español)
                  Me.txt_cantidad_general = IIf(IsNull(rs!FLOA_EMO_CANTIDAD_TRANSFORMAR), 1, rs!FLOA_EMO_CANTIDAD_TRANSFORMAR)
                  Me.txt_cantidad_general.Enabled = False
                  Me.lv_salidas.ListItems.Clear
                  While Not rsaux3.EOF
                       Set list_item = lv_salidas.ListItems.Add(, , rsaux3!vcha_Art_Articulo_id)
                       rsaux2.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + rsaux3!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                       If Not rsaux2.EOF Then
                          list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_art_nombre_Español), "", rsaux2!vcha_art_nombre_Español)
                       End If
                       rsaux2.Close
                       list_item.SubItems(2) = IIf(IsNull(rsaux3!floa_ent_Cantidad), 0, rsaux3!floa_ent_Cantidad)
                       rsaux3.MoveNext
                  Wend
                  rsaux3.Close
                  Me.txt_almacen.Enabled = False
                  Me.txt_nombre_almacen.Enabled = False
                  Me.txt_codigo_general.Enabled = False
                  Me.txt_nombre_articulo_general.Enabled = False
                  var_primera_vez = False
                  var_estatus_movimiento = IIf(IsNull(rs!char_Emo_estatus), "", rs!char_Emo_estatus)
                  If var_estatus_movimiento <> "I" Then
                     Me.txt_codigo.Enabled = True
                     Me.txt_codigo.SetFocus
                  Else
                     Me.txt_codigo.Enabled = False
                     Me.txt_cantidad.Enabled = False
                     Me.txt_foco.Enabled = False
                  End If
               Else
                  MsgBox "Código de artículo general no existe", vbOKOnly, "ATENCION"
               End If
               rsaux1.Close
            Else
               MsgBox "El almacén del movimiento no existe", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de folio incorrecto", vbOKOnly, "ATENCION"
      End If
      Me.frm_busqueda.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
    Me.frm_busqueda.Visible = False
End Sub


Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.txt_cantidad_eliminar) <= CDbl(Me.lv_salidas.selectedItem.SubItems(2)) Then
            rsaux.Open "UPDATE TB_TEMPORAL_ENTRADAS SET FLOA_ENT_CANTIDAD = ISNULL(FLOA_ENT_CANTIDAD,0) - " + txt_cantidad_eliminar + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_ENT_NUMERO = " + CStr(Me.txt_folio) + " AND VCHA_ART_ARTICULO_ID= '" + lv_salidas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_salidas.selectedItem.SubItems(2) = CDbl(Me.lv_salidas.selectedItem.SubItems(2)) - CDbl(Me.txt_cantidad_eliminar)
         Else
            MsgBox "La cantidad no debe de ser superior a la cantidad en la salida", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad eliminar incorrecta", vbOKOnly, "ATENCION"
      End If
      Me.lv_salidas.SetFocus
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_cantidad_general_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.Enabled = True
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_general_LostFocus()
   If Not IsNumeric(Me.txt_cantidad_general) Then
      Me.txt_cantidad_general = "1"
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   Me.txt_cantidad = 1
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad) Then
         Me.txt_foco.Enabled = True
         Me.txt_foco.SetFocus
      End If
   End If
End Sub

Private Sub txt_codigo_general_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_cantidad_general.Enabled = True
      Me.txt_cantidad_general.SetFocus
   End If
End Sub

Private Sub txt_codigo_general_LostFocus()
   If Trim(Me.txt_codigo_general) <> "" Then
      rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_articulo_general = IIf(IsNull(rs!vcha_art_nombre_Español), "", rs!vcha_art_nombre_Español)
         Me.txt_codigo_general.Enabled = False
         Me.txt_nombre_articulo_general.Enabled = False
      Else
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_codigo_general = IIf(IsNull(rsaux1!vcha_Art_Articulo_id), "", rsaux1!vcha_Art_Articulo_id)
               Me.txt_nombre_articulo_general = IIf(IsNull(rsaux1!vcha_art_nombre_Español), "", rsaux1!vcha_art_nombre_Español)
               Me.txt_codigo_general.Enabled = False
               Me.txt_nombre_articulo_general.Enabled = False
            Else
               MsgBox "Código de artículo incorrecto", vbOKOnly, "ATENCION"
            End If
            rsaux1.Close
         Else
            MsgBox "Código de artículo incorrecto", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
      rs.Close
   Else
      Me.txt_nombre_articulo_general = ""
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.txt_codigo <> "" Then
         If Me.txt_codigo_general <> Me.txt_codigo Then
            rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_cantidad = ""
               Me.txt_cantidad.SetFocus
            Else
               rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rsaux1.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.txt_codigo = rsaux1!vcha_Art_Articulo_id
                  rsaux1.Close
                  Me.txt_cantidad = ""
                  Me.txt_cantidad.SetFocus
               Else
                  MsgBox "Código de artículo incorrecto", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            End If
            rs.Close
         Else
            MsgBox "El código no puede ser el mismo", vbOKOnly, "ATENCION"
            Me.txt_codigo = ""
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Dim var_inserta As Boolean
   var_almacen_Destino = txt_almacen
   If Trim(txt_codigo.Text) <> "" Then
      If IsNumeric(Me.txt_cantidad) Then
         var_cantidad_leida = CDbl(Me.txt_cantidad)
         var_almacen_Destino = CStr(Me.txt_almacen)
         bandera_suma = False
         If var_primera_vez = True Then
            var_inserta = False
            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, txt_almacen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_Destino, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", Me.txt_codigo_general, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
            var_numero_folio = var_numero_folio_regreso
            txt_folio = var_numero_folio
            If Not IsNumeric(Me.txt_cantidad_general) Then
               Me.txt_cantidad_general = 1
            End If
            rsaux.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET FLOA_EMO_CANTIDAD_TRANSFORMAR = " + CStr(CDbl(Me.txt_cantidad_general)) + ", vcha_emo_descripcion_transformar = '" + Mid(Me.txt_nombre_articulo_general, 1, 50) + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            var_primera_vez = False
         End If
         var_posible_leido = 1
         If var_posible_leido = 1 Then
            If var_costo = 0 Then
               rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + txt_almacen + "' and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_costo = IIf(IsNull(rs!FLOA_eXI_COSTO), 0, rs!FLOA_eXI_COSTO)
               Else
                  var_costo = 0
               End If
               rs.Close
            End If
            rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_descripcion_articulo = IIf(IsNull(rs!vcha_art_nombre_Español), "", rs!vcha_art_nombre_Español)
            Else
               var_descripcion_articulo = ""
            End If
            rs.Close
            Cadena = "select * from TB_TEMPORAL_entradas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + txt_almacen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         
            If Not rs.EOF Then
               var_inserta = False
               rsaux.Open "uPdate tb_temporal_entradas set floa_ent_Cantidad = ISNULL(floa_ent_cantidad,0) + " + CStr(CDbl(Me.txt_cantidad)) + "  where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + txt_almacen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               rs.Close
               valor = Trim(txt_codigo)
               Set itmfound = lv_salidas.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
               lv_salidas.selectedItem.SubItems(2) = lv_salidas.selectedItem.SubItems(2) + var_cantidad_leida
               var_renglon = lv_salidas.selectedItem.Index
               txt_total = CStr(CDbl(txt_total) + var_cantidad_leida)
               'Call ilumina_grid
            Else
               var_inserta = False
               rsaux.Open "INSERT INTO TB_TEMPORAL_entradas (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + txt_almacen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.Close
               Set list_item = Me.lv_salidas.ListItems.Add(, , Trim(txt_codigo))
               list_item.SubItems(1) = var_descripcion_articulo
               list_item.SubItems(2) = var_cantidad_leida
               var_renglon = lv_salidas.ListItems.Count
               txt_total = CStr(CDbl(txt_total) + var_cantidad_leida)
               'Call ilumina_grid
            End If
         Else
            frmmensaje.lbl_mensaje = var_kanban_mensaje
            frmmensaje.Show 1
            txt_codigo = ""
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
      Me.txt_cantidad = ""
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call pro_enfoque(KeyAscii)
    Else
       KeyAscii = 0
    End If
End Sub

Private Sub txt_nombre_articulo_general_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Else
      KeyAscii = 0
   End If
End Sub
