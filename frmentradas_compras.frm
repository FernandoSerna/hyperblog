VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentradas_compras 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmentradas_compras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7635
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7155
      Picture         =   "frmentradas_compras.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   735
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2625
      Left            =   1530
      TabIndex        =   35
      Top             =   900
      Width           =   5820
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2145
         Left            =   30
         TabIndex        =   36
         Top             =   420
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   3784
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
            Text            =   "Nombre"
            Object.Width           =   10107
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   37
         Top             =   120
         Width           =   5745
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1560
      Index           =   0
      Left            =   5310
      TabIndex        =   15
      Top             =   1125
      Width           =   2205
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
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   645
         Width           =   2025
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   2130
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8130
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2910
      Width           =   1125
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmentradas_compras.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmentradas_compras.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Buscar Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmentradas_compras.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmentradas_compras.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   735
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1185
      TabIndex        =   0
      Top             =   1035
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   120
         Width           =   3060
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   15
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
            Picture         =   "frmentradas_compras.frx":130C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":1BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":24C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":2A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":3338
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":3C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":44EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":45FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":4710
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":4822
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":4934
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_compras.frx":4A46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   585
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   18
      Top             =   975
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      Height          =   1560
      Index           =   1
      Left            =   105
      TabIndex        =   9
      Top             =   1125
      Width           =   5190
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   990
         TabIndex        =   33
         Top             =   1155
         Width           =   1695
      End
      Begin VB.TextBox txt_nombre_proveedor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   810
         Width           =   3300
      End
      Begin VB.TextBox txt_nombre_almacen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   465
         Width           =   3300
      End
      Begin VB.TextBox txt_clave_almacen 
         Height          =   315
         Left            =   990
         TabIndex        =   30
         Top             =   465
         Width           =   810
      End
      Begin VB.TextBox txt_clave_proveedor 
         Height          =   315
         Left            =   990
         TabIndex        =   10
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   34
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   510
         Width           =   585
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   5115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   855
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4620
      Left            =   105
      TabIndex        =   19
      Top             =   2625
      Width           =   7425
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
         Left            =   5115
         TabIndex        =   24
         Top             =   555
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   21
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            MaxLength       =   10
            TabIndex        =   22
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
            TabIndex        =   23
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
         Left            =   1545
         TabIndex        =   20
         Top             =   495
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3435
         Left            =   45
         TabIndex        =   25
         Top             =   1110
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   6059
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
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8617
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   28
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   675
         Width           =   1395
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
      Left            =   90
      TabIndex        =   29
      Top             =   90
      Width           =   7335
   End
End
Attribute VB_Name = "frmentradas_compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_año As Integer
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
Dim var_tipo_lista As Integer

Private Sub cmd_buscar_Click()
   var_ventana = 1
   frm_busqueda.Visible = True
   txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_cancelar_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Then
         MsgBox "El Movimiento ya fue cancelado", vbOKOnly, "ATENCION"
      Else
         If var_estatus_movimiento = "I" Then
            If var_fecha_movimiento <> Date Then
               var_posible_accion = False
               frmsupervisor1.Show
               If var_posible_accion = True Then
                  si = MsgBox("¿Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la cancelación del movimiento", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ENC_MOV_CANCELACION = New TB_ENC_MOV_CANCELACION
                        var_actualizar = False
                        var_actualizar = TB_ENC_MOV_CANCELACION.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "C", var_global_supervisor_1, var_global_supervisor_2)
                        MsgBox "El movimiento a sido cancelado", vbOKOnly, "ATENCION"
                        var_estatus_movimiento = "C"
                     End If
                  End If
               End If
            Else
               var_posible_accion = False
               frmsupervisor1.Show
               If var_posible_accion = True Then
                  si = MsgBox("¿Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la cancelación del movimiento", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ENC_MOV_CANCELACION = New TB_ENC_MOV_CANCELACION
                        var_actualizar = False
                        var_actualizar = TB_ENC_MOV_CANCELACION.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "C", var_global_supervisor_1, var_global_supervisor_2)
                        var_estatus_movimiento = "C"
                        MsgBox "El movimiento a sido cancelado", vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            End If
         Else
            MsgBox "El Movimiento no a sido cerrado aun", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
   End If
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
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_ENTRADAS_compra.rpt")
         reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPRA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENTRADAS_COMPRA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + txt_clave_almacen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      Else
         var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            cnn.BeginTrans
            Cadena = "select * from tb_temporal_entradas where vcha_alm_almacen_id = '" + txt_clave_almacen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_inserta = False
                  rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, inte_ent_año) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(rs!floa_ent_Cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            var_estatus_movimiento = "I"
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, txt_clave_almacen, var_clave_movimiento, var_numero_folio, "I", Now, 1)
            cnn.CommitTrans
            Set reporte = appl.OpenReport(App.Path + "\rep_ENTRADAS_compra.rpt")
            reporte.RecordSelectionFormula = "{VW_ENTRADAS_COMPRA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_ENTRADAS_COMPRA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
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
            rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + txt_clave_almacen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_clave_almacen = ""
   txt_nombre_almacen = ""
   txt_clave_proveedor = ""
   txt_nombre_proveedor = ""
   txt_factura = ""
   txt_folio = ""
   txt_codigo = ""
   txt_cantidad = ""
   txt_cantidad_eliminar = ""
   frm_lista.Visible = False
   lv_entradas.ListItems.Clear
   txt_clave_almacen.Enabled = True
   txt_clave_almacen.SetFocus
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

Private Sub Form_Load()
   var_numero_folio = 0
   var_año = 2005
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
   txt_clave_almacen.Enabled = False
   txt_clave_proveedor.Enabled = False
   var_ventana = 0
   var_estatus_movimiento = ""
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_factura.Enabled = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_entradas_compras)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         If var_causa_devolucion = True Then
            rs.Open "select * from tb_causas_devolucion order by vcha_cde_nombre", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_elimina = True
               lv_causas_devolucion.ListItems.Clear
               While Not rs.EOF
                  Set list_item = lv_causas_devolucion.ListItems.Add(, , rs!INTE_CDE_CAUSA_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                  rs.MoveNext
               Wend
               rs.Close
               lv_causas_devolucion.SetFocus
            Else
               var_elimina = False
               var_ventana = 1
               frm_eliminar.Visible = True
               txt_cantidad_eliminar.SetFocus
            End If
         Else
            var_elimina = False
            var_ventana = 1
            frm_eliminar.Visible = True
            txt_cantidad_eliminar.SetFocus
         End If
      End If
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         txt_clave_almacen = lv_lista.selectedItem.SubItems(1)
         txt_nombre_almacen = lv_lista.selectedItem
         txt_clave_almacen.Enabled = False
         txt_clave_proveedor.Enabled = True
         txt_clave_proveedor.SetFocus
         frm_lista.Visible = False
      End If
      If var_tipo_lista = 2 Then
         txt_clave_proveedor = lv_lista.selectedItem.SubItems(1)
         txt_nombre_proveedor = lv_lista.selectedItem
         txt_clave_proveedor.Enabled = False
         txt_factura.Enabled = True
         txt_factura.SetFocus
         frm_lista.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
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
               txt_clave_almacen = rs!VCHA_ALM_ALMACEN_ID
               var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
               txt_factura = rs!VCHA_EMO_FACTURA
               txt_clave_proveedor = rs!VCHA_PRO_PROVEEDOR_ID
               rsaux.Open "SELECT * FROM TB_PROVEEDORES WHERE VCHA_PRO_PROVEEDOR_ID = '" + txt_clave_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_nombre_proveedor = IIf(IsNull(rsaux!VCHA_PRO_NOMBRE), "", rsaux!VCHA_PRO_NOMBRE)
               End If
               rsaux.Close
               txt_factura.Enabled = False
               lv_entradas.ListItems.Clear
               var_primera_vez = False
               var_numero_folio = rs!INTE_EMO_NUMERO
               txt_folio = var_numero_folio
               rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_almacen = rsaux!VCHA_ALM_NOMBRE
               rsaux.Close
               rsaux.Open "select * from tb_temporal_entradas where inte_ent_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                     rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Set list_item = lv_entradas.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                        list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                        list_item.SubItems(2) = IIf(IsNull(rsaux!floa_ent_Cantidad), "", rsaux!floa_ent_Cantidad)
                        rsaux2.Close
                        rsaux.MoveNext:
                     End If
                  Wend
               End If
               rsaux.Close
               If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                  txt_codigo.Enabled = False
                  txt_cantidad.Visible = False
                  lbl_cantidad.Visible = False
                  txt_foco.Enabled = False
               Else
                  txt_foco.Enabled = False
                  txt_codigo.Enabled = True
                  txt_cantidad.Visible = False
                  lbl_cantidad.Visible = False
               End If
            Else
               MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
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
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_eliminar) Then
         Dim var_posible_eliminar As Boolean
         Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
         var_cantidad_eliminar = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If var_cantidad_eliminar >= (lv_entradas.selectedItem.SubItems(2) * 1) Then
            var_posible_eliminar = False
         End If
         If var_posible_eliminar = True Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, txt_clave_almacen, var_clave_movimiento, var_numero_folio, lv_entradas.selectedItem, 0 - Val(txt_cantidad_eliminar))
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
         Else
            MsgBox "La cantidad a eliminar supera a la posible a eliminar", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
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

Private Sub txt_clave_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_lista = 1
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select vcha_alm_nombre, vcha_alm_almacen_id from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         numero_items = 0
         While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext:
            numero_items = numero_items + 1
         Wend
         rs.Close
         If numero_items > 8 Then
            lv_lista.ColumnHeaders(1).Width = 5430
         Else
            lv_lista.ColumnHeaders(1).Width = 5630
         End If
         lbl_lista = "Lista de Proveedores"
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         rs.Open "select  vcha_alm_nombre, vcha_alm_almacen_idfrom vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         numero_items = 0
         While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext:
            numero_items = numero_items + 1
         Wend
         rs.Close
         If numero_items > 8 Then
            lv_lista.ColumnHeaders(1).Width = 5430
         Else
            lv_lista.ColumnHeaders(1).Width = 5630
         End If
         lbl_lista = "Lista de Almacenes"
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_almacen) <> "" Then
         If var_tipo_permiso = 1 Then
            rs.Open "select vcha_alm_nombre, vcha_alm_almacen_id from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
               txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
               txt_clave_almacen.Enabled = False
               txt_clave_proveedor.Enabled = True
               txt_clave_proveedor.SetFocus
            Else
               MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
               txt_clave_almacen = ""
               txt_nombre_almacen = ""
            End If
            rs.Close
         Else
            rs.Open "select  vcha_alm_nombre, vcha_alm_almacen_id from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
               txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
               txt_clave_proveedor.Enabled = True
               txt_clave_proveedor.SetFocus
            Else
               MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
               txt_clave_almacen = ""
               txt_nombre_almacen = ""
            End If
            rs.Close
         End If
      End If
   End If
End Sub

Private Sub txt_clave_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_clave_proveedor_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_lista = 2
      lv_lista.ListItems.Clear
      rs.Open "select vcha_pro_nombre, vcha_pro_proveedor_id from tb_proveedores order by vcha_pro_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      numero_items = 0
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext:
            numero_items = numero_items + 1
      Wend
      rs.Close
      If numero_items > 8 Then
         lv_lista.ColumnHeaders(1).Width = 5430
      Else
         lv_lista.ColumnHeaders(1).Width = 5630
      End If
      lbl_lista = "Lista de Proveedores"
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_proveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_proveedor) <> "" Then
         rs.Open "select * from tb_proveedores where vcha_pro_proveedor_id ='" + txt_clave_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_proveedor = rs!VCHA_PRO_NOMBRE
            txt_factura.Enabled = True
            txt_clave_proveedor.Enabled = False
            txt_factura.SetFocus
         Else
            MsgBox "Clave de proveedor incorrecta", vbOKOnly, "ATENCION"
            txt_clave_proveedor = ""
            txt_nombre_proveedor = ""
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_clave_proveedor_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_codigo_seleccionado = ""
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Dim var_recontable As Integer
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      var_verificador = True
      If Len(Trim(txt_codigo)) = 12 Then
         Call calcula_verificador(Trim(txt_codigo))
      End If
      If var_verificador = True Then
         var_caja = Left(txt_codigo, 6)
         If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000011" Or var_caja = "0000012" Or var_caja = "0000013" Or var_caja = "0000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
            var_cantidad_caja = CInt(var_caja)
            txt_codigo = Mid(txt_codigo, 7, 5)
         End If
         var_costo = 0
         var_precio = 0
         If Trim(txt_codigo) <> "" Then
            rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "select * from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  If IsNull(rs!inte_Art_salida_masiva) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs!inte_Art_salida_masiva
                  End If
                  var_descripcion_articulo = rs!vcha_art_nombre_español
                  var_costo = rsaux!floa_cpr_costo_predeterminado
                  var_precio = rs!mone_art_precio_base
                  rs.Close
                  If var_recontable = 1 Then
                     var_cantidad_leida = 1#
                     lbl_cantidad.Visible = True
                     txt_cantidad.Visible = True
                     txt_cantidad.SetFocus
                  Else
                     var_cantidad_leida = 1#
                     txt_foco.Enabled = True
                     txt_foco.SetFocus
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "Artículo invalido para el proveedor seleccionado"
                  frmmensaje.Show
                  'MsgBox "Artículo invalido para el proveedor seleccionado", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            Else
               rs.Close
               rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  txt_codigo = rs(0).Value
                  rs.Close
                  rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "select * from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        If var_cantidad_caja = 0 Then
                           If IsNull(rs!inte_Art_salida_masiva) Then
                              var_recontable = 0
                           Else
                              var_recontable = rs!inte_Art_salida_masiva
                           End If
                        Else
                           var_recontable = 0
                        End If
                        var_descripcion_articulo = rs!vcha_art_nombre_español
                        var_costo = rsaux!floa_cpr_costo_predeterminado
                        var_precio = rs!mone_art_precio_base
                        rs.Close
                        If var_recontable = 1 Then
                           var_cantidad_leida = 1#
                           lbl_cantidad.Visible = True
                           txt_cantidad.Visible = True
                           txt_cantidad.SetFocus
                        Else
                           If var_cantidad_caja = 0 Then
                              var_cantidad_leida = 1#
                           Else
                              var_cantidad_leida = var_cantidad_caja
                           End If
                           txt_foco.Enabled = True
                           txt_foco.SetFocus
                        End If
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "Artículo invalido para el proveedor seleccionado"
                        frmmensaje.Show
                        'MsgBox "Artículo invalido para el proveedor seleccionado", vbOKOnly, "ATENCION"
                     End If
                     rsaux.Close
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show
                     'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El artículo no existe"
                  frmmensaje.Show
                  'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  rs.Close
               End If
            End If
         Else
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "Error en el Código"
         frmmensaje.Show
         MsgBox "Error en el Código", vbOKOnly, "ATENCION"
      End If
      If rs.State = 1 Then
         rs.Close
      End If
   End If
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_almacen) <> "" Then
         If Trim(txt_clave_proveedor) <> "" Then
            If Trim(txt_factura) <> "" Then
               txt_codigo.Enabled = True
               txt_factura.Enabled = False
               txt_codigo.SetFocus
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      bandera_suma = False
      If var_primera_vez = True Then
         var_inserta = False
         var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, txt_clave_almacen, var_clave_movimiento, Now, var_numero_folio, 0, "", txt_clave_proveedor, "", txt_clave_almacen, "", var_clave_usuario_global, fun_NombrePc, txt_factura, "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
         var_numero_folio = var_numero_folio_regreso
         txt_folio = var_numero_folio
         var_primera_vez = False
      End If
      Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = " + txt_clave_almacen + "and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = " + txt_codigo
      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_inserta = False
         var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, txt_clave_almacen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_año)
         rs.Close
         valor = Trim(txt_codigo)
         Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
         itmfound.EnsureVisible
         itmfound.Selected = True
         lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) + var_cantidad_leida
      Else
         var_inserta = False
         var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, txt_clave_almacen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
         rs.Close
         Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
         list_item.SubItems(1) = var_descripcion_articulo
         list_item.SubItems(2) = var_cantidad_leida
      End If
      txt_codigo.SetFocus
   End If
End Sub
