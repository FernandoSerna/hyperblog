VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_reempaque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salidas para reempaque"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7605
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1335
      TabIndex        =   33
      Top             =   525
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   34
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
         TabIndex        =   35
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   915
      TabIndex        =   21
      Top             =   1035
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Index           =   0
      Left            =   5175
      TabIndex        =   18
      Top             =   1095
      Width           =   2295
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
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
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   2160
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4860
      Left            =   45
      TabIndex        =   8
      Top             =   2430
      Width           =   7425
      Begin VB.TextBox txt_cantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
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
         TabIndex        =   13
         Top             =   495
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   10
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   11
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
            TabIndex        =   12
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
         TabIndex        =   9
         Top             =   450
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_traspasossalidas 
         Height          =   3360
         Left            =   45
         TabIndex        =   14
         Top             =   1050
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   5927
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   37
         Top             =   4425
         Width           =   2715
      End
      Begin VB.Label lbl_cantidad_leida 
         Alignment       =   1  'Right Justify
         Caption         =   "9999999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4395
         TabIndex        =   36
         Top             =   4425
         Width           =   2715
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   17
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   570
      Width           =   7455
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   7845
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3615
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7125
      Picture         =   "frmsalidas_reempaque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmsalidas_reempaque.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmsalidas_reempaque.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmsalidas_reempaque.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmsalidas_reempaque.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   735
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
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
            Picture         =   "frmsalidas_reempaque.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_reempaque.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   600
      Top             =   0
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Index           =   1
      Left            =   45
      TabIndex        =   24
      Top             =   1095
      Width           =   5100
      Begin VB.TextBox txt_nombre_almacen_destino 
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   795
         Width           =   3270
      End
      Begin VB.TextBox txt_nombre_almacen_origen 
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   450
         Width           =   3270
      End
      Begin VB.TextBox txt_clave_almacen_destino 
         Height          =   315
         Left            =   780
         TabIndex        =   30
         Top             =   795
         Width           =   990
      End
      Begin VB.TextBox txt_clave_almacen_origen 
         Height          =   315
         Left            =   780
         TabIndex        =   29
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   510
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   5025
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   975
      Width           =   7455
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      TabIndex        =   28
      Top             =   75
      Width           =   7335
   End
End
Attribute VB_Name = "frmsalidas_reempaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_kanban As String
Dim var_almacen_Destino As String
Dim var_almacen_origen As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_almacen As String
Dim var_correo_electronico As String
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_tipo_lista As Integer
Dim var_renglon As Double
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


Sub ilumina_grid()
   var_n = lv_traspasossalidas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_traspasossalidas.ListItems.Item(var_i).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_traspasossalidas.ListItems.Item(var_i).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_traspasossalidas.ListItems.Item(var_renglon).Selected = True
      lv_traspasossalidas.selectedItem.EnsureVisible
   End If
   lv_traspasossalidas.Refresh
End Sub


Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function






Private Sub cmd_buscar_Click()
            frm_busqueda.Visible = True: var_ventana = 1
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
   'On Error GoTo salir:
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Dim var_correo_electronico As String
   Dim var_Archivo As String
   If Dir(var_ruta & "\reempaque.dbf") <> "" Then
      If var_numero_folio > 0 Then
         If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_salidas_reempaque.rpt")
            reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            If var_tipo_almacen = "T" Then
               Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
            End If
            Set var_tabla = CreateObject("ADODB.connection")
            'var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
             'var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & App.Path & ";DefaultDir=" & var_ruta & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
             var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"

            Cadena = "select * from tb_entradas where vcha_alm_almacen_id = " + var_almacen_Destino + " and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            var_Archivo = Trim(CStr(var_numero_folio))
            var_n = Len(Trim(var_Archivo))
            While var_n < 8
                  var_Archivo = "0" + var_Archivo
                  var_n = var_n + 1
            Wend
            var_Archivo = var_Archivo
            rsaux2.Open "delete from reempaque", var_tabla, adOpenDynamic, adLockOptimistic
            If Dir(var_ruta & "\" + Trim(var_Archivo) + ".dbf") <> "" Then
               Kill var_ruta & "\" + Trim(var_Archivo) + ".dbf"
            End If
            var_copia = CopyFile(var_ruta & "\reempaque.dbf", var_ruta & "\" + var_Archivo + ".dbf", 1)
            
            Cadena = "select * from tb_Salidas where VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_codigo = rs!vcha_Art_articulo_id
                  rsaux3.Open "select vcha_art_codigo_externo from tb_Articulos WHERE VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  'If Not rsaux3.EOF Then
                  '   var_codigo = rsaux3!VCHA_aRT_CODIGO_EXTERNO
                  'End If
                  rsaux3.Close
                  var_almacen_Destino_2 = var_almacen_Destino
                  If var_almacen_Destino = "6" Then
                     var_almacen_Destino_2 = "4"
                  End If
                  var_codigo = Mid(rs!vcha_Art_articulo_id, 7, 5)
                  rsaux2.Open "insert into " + var_Archivo + " (numnota,planta,codigo,descripcio,tallas,talla1,talla2,talla3,talla4,talla5,talla6,costo,cant1,cant2,cant3,cant4,cant5,cant6,anocosto) values ('" + Str(var_numero_folio) + "', '" + var_almacen_Destino_2 + "', '" + var_codigo + "',' ', 1,0,0,0,0,0,0," + Str(rs!floa_Sal_costo) + ", " + Str(rs!FLOA_sAL_cANTIDAD) + ",0,0,0,0,0,'" + CStr(rs!INTE_sAL_AÑO) + "')", var_tabla, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            var_tabla.Close
            Set var_tabla = Nothing
            Cadena = "select * from tb_entradas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            var_archivo_origen = var_ruta & "\" + var_Archivo + ".dbf"
            var_archivo_destino = var_ruta & "\" + var_almacen_origen + var_Archivo + ".dbf"
            var_eliminar = DeleteFile(var_ruta & "\" + Trim(var_almacen_origen) + var_Archivo + ".dbf")
            var_copia = CopyFile(var_archivo_origen, var_archivo_destino, 1)
            var_Archivo = var_almacen_origen + var_Archivo
            rs.Close
            rs.Open "select vcha_alm_correo from tb_almacenes where  vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_correo_electronico = IIf(IsNull(rs!vcha_alm_correo), "", rs!vcha_alm_correo)
            Else
               var_correo_electronico = ""
            End If
            rs.Close
            If var_correo_electronico <> "" Then
               If MAPISession1.SessionID = 0 Then
                  MAPISession1.SignOn
               End If
               MAPIMessages1.SessionID = MAPISession1.SessionID
               MAPIMessages1.Compose
               MAPIMessages1.RecipDisplayName = var_correo_electronico
               MAPIMessages1.RecipAddress = var_correo_electronico
               MAPIMessages1.MsgSubject = "Nota de reempaque " + Trim(var_Archivo)
               MAPIMessages1.MsgNoteText = "Se adjunta archivo de reempaque"
               MAPIMessages1.AttachmentPathName = var_ruta & "\" + var_Archivo + ".dbf"
               MAPIMessages1.Send True
               
               If MAPISession1.SessionID > 0 Then
                  MAPISession1.SignOff
               End If
            Else
               MsgBox "No se a indicado una dirección de correo electrónico", vbOKOnly, "ATENCION"
            End If
         Else
            var_posible_Cantidad = 1
            If var_empresa = "18" Or var_empresa = "31" Then
               Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and floa_Sal_cantidad > 0"
               rsaux10.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux10.EOF
                     rsaux9.Open "select * from tb_existencias where vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_Art_Articulo_id = '" + rsaux10!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        var_cantidad = IIf(IsNull(rsaux9!floa_Exi_Cantidad_disponible), 0, rsaux9!floa_Exi_Cantidad_disponible)
                        If var_empresa = "18" Then
                           If rsaux10!vcha_Art_articulo_id = "360010000002" Or rsaux10!vcha_Art_articulo_id = "360020000009" Or rsaux10!vcha_Art_articulo_id = "900000000003" Or rsaux10!vcha_Art_articulo_id = "911110000005" Then
                              var_cantidad = Round(IIf(IsNull(rsaux10!FLOA_sAL_cANTIDAD), 0, rsaux10!FLOA_sAL_cANTIDAD), 4) + 1
                           End If
                        End If
                        
                        If Round(var_cantidad, 4) < Round(IIf(IsNull(rsaux10!FLOA_sAL_cANTIDAD), 0, rsaux10!FLOA_sAL_cANTIDAD), 4) Then
                           var_posible_Cantidad = 0
                           If var_cadena_articulos = "" Then
                              rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux8.EOF Then
                                 var_nombre_articulo = IIf(IsNull(rsaux8!vcha_art_nombre_español), "", rsaux8!vcha_art_nombre_español)
                              Else
                                 var_nombre_articulo = ""
                              End If
                              rsaux8.Close
                              var_cadena_articulos = rsaux10!vcha_Art_articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!FLOA_sAL_cANTIDAD) + "]"
                           Else
                              rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux8.EOF Then
                                 var_nombre_articulo = IIf(IsNull(rsaux8!vcha_art_nombre_español), "", rsaux8!vcha_art_nombre_español)
                              Else
                                 var_nombre_articulo = ""
                              End If
                              rsaux8.Close
                              var_cadena_articulos = var_cadena_articulos + ", " + rsaux10!vcha_Art_articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!FLOA_sAL_cANTIDAD) + "]"
                           End If
                        
                        
                        End If
                     Else
                        If var_cadena_articulos = "" Then
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_art_nombre_español), "", rsaux8!vcha_art_nombre_español)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = rsaux10!vcha_Art_articulo_id + " " + var_nombre_articulo
                        Else
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_art_nombre_español), "", rsaux8!vcha_art_nombre_español)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = var_cadena_articulos + ", " + rsaux10!vcha_Art_articulo_id + " " + var_nombre_articulo
                        End If
                        var_posible_Cantidad = 0
                     End If
                     rsaux9.Close
                     rsaux10.MoveNext
               Wend
               rsaux10.Close
            End If
            If var_posible_Cantidad = 1 Then
               var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
               If var_si = 1 Then
                  Set var_tabla = CreateObject("ADODB.connection")
                  VAR_MAQUINA = fun_NombrePc
                  If Not UCase(VAR_MAQUINA) = "JFSERNA" Then
                     var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                  Else
                     var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & var_ruta & ";DefaultDir=" & var_ruta & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
                  End If
                  var_Archivo = Trim(CStr(var_numero_folio))
                  var_n = Len(Trim(var_Archivo))
                  While var_n < 8
                        var_Archivo = "0" + var_Archivo
                        var_n = var_n + 1
                  Wend
                  var_Archivo = var_Archivo
                  Cadena = "select * from tb_temporal_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                  cnn.BeginTrans
                  
                  var_posible_cerrar_KANBAN = True
                  If var_posible_kanban = 1 Then
                     Set TB_PROC_KANBANS_EN_MOVIMIENTO = New TB_PROC_KANBANS_EN_MOVIMIENTO
                     var_inserta = TB_PROC_KANBANS_EN_MOVIMIENTO.Anadir(Me.txt_clave_almacen_origen, var_clave_movimiento, CDbl(Me.txt_folio), "", "")
                     If var_kanban_exito = "N" Then
                        var_posible_cerrar_KANBAN = False
                     End If
                  Else
                     var_posible_cerrar_KANBAN = True
                  End If
                  If var_posible_cerrar_KANBAN = True Then
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                           var_suma_cantidad = 0
                           var_cantidad_llegar = IIf(IsNull(rs!FLOA_sAL_cANTIDAD), 0, rs!FLOA_sAL_cANTIDAD)
                           var_cantidad = 0
                           While var_suma_cantidad < var_cantidad_llegar
                                 rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!vcha_Art_articulo_id + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                                       var_año = 2004
                                       var_suma_cantidad = var_cantidad_llegar
                                       var_cantidad = var_cantidad_llegar
                                       var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                    Else
                                       var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                       If var_cantidad_disponible > 0 Then
                                          var_año = 2004
                                          var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                          var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                          var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                       Else
                                          var_año = 2005
                                          var_cantidad = rs!FLOA_sAL_cANTIDAD - var_suma_cantidad
                                          var_suma_cantidad = var_cantidad_llegar
                                          var_costo = IIf(IsNull(rsaux2!floa_exi_costo_2005), 0, rsaux2!floa_exi_costo_2005)
                                          If var_costo = 0 Then
                                             var_costo = IIf(IsNull(rsaux2!FLOA_eXI_COSTO), 0, rsaux2!FLOA_eXI_COSTO)
                                          End If
                                          
                                       End If
                                    End If
                                 Else
                                    var_año = 2005
                                    var_suma_cantidad = var_cantidad_llegar
                                    var_cantidad = var_cantidad_llegar
                                    rsaux4.Open "select * from tb_existencias where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "' and vcha_alm_almacen_id = '8'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_costo = IIf(IsNull(rsaux4!floa_exi_costo_2005), 0, rsaux4!floa_exi_costo_2005)
                                    Else
                                       var_costo = 0
                                    End If
                                    rsaux4.Close
                                    If var_costo = 0 Then
                                       rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux4.EOF Then
                                          var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                                       Else
                                          var_costo = 0
                                       End If
                                       rsaux4.Close
                                    End If
                                 End If
                                 rsaux2.Close
                                 rsaux4.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_SAL_AÑO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + ", " + CStr(rs!floa_Sal_precio) + ",0, " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux4.Open "INSERT INTO TB_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_AÑO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + ", " + CStr(rs!floa_Sal_precio) + ",'" + var_almacen_origen + "'," + CStr(var_año) + ")", cnn, adOpenDynamic, adLockBatchOptimistic
                                 rsaux.Open "INSERT INTO TB_REEMPAQUE_SALIDA (VCHA_REE_FOLIO, VCHA_REE_ALMACEN_ORIGEN, VCHA_REE_ALMACEN_DESTINO,INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_REE_CANTIDAD_SALIDA ,FLOA_REE_COSTO_SALIDA, FLOA_REE_CANTIDAD_LEIDA, INTE_REE_AÑO) VALUES ('" + Trim(var_almacen_origen) + Trim(var_Archivo) + "', '" + var_almacen_origen + "', '" + var_almacen_Destino + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + ", 0, " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                           Wend
                           rs.MoveNext
                     Wend
                     rs.Close
                     var_estatus_movimiento = "I"
                     rs.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I', dtim_emo_fecha_finalizo = getdate() where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                     txt_codigo.Enabled = False
                     txt_foco.Enabled = False
                  End If
                  cnn.CommitTrans
                  If var_posible_cerrar_KANBAN = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_salidas_reempaque.rpt")
                     reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                     If var_tipo_almacen = "T" Then
                         Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
                     End If
                     var_archivo_origen = var_ruta & "\" + var_Archivo + ".dbf"
                     var_archivo_destino = var_ruta & "\" + var_almacen_origen + var_Archivo + ".dbf"
                     If Dir(var_ruta + "\" + Trim(var_Archivo) + ".dbf") <> "" Then
                        Kill var_ruta + "\" + Trim(var_Archivo) + ".dbf"
                     End If
                     If Dir(Trim(var_archivo_destino)) <> "" Then
                        Kill Trim(var_archivo_destino)
                     End If
                     var_copia = CopyFile(var_ruta & "\reempaque.dbf", var_ruta & "\" + var_Archivo + ".dbf", 1)
                     Cadena = "select * from tb_Salidas where VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                           var_codigo = rs!vcha_Art_articulo_id
                           rsaux3.Open "select vcha_art_codigo_externo from tb_Articulos WHERE VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              var_codigo = rsaux3!VCHA_aRT_CODIGO_EXTERNO
                           End If
                           rsaux3.Close
                           var_almacen_Destino_2 = var_almacen_Destino
                           If var_almacen_Destino_2 = "6" Then
                              var_almacen_Destino_2 = "4"
                           End If
                           rsaux2.Open "insert into " + var_Archivo + " (numnota,planta,codigo,descripcio,tallas,talla1,talla2,talla3,talla4,talla5,talla6,costo,cant1,cant2,cant3,cant4,cant5,cant6,anocosto) values ('" + Str(var_numero_folio) + "', '" + var_almacen_Destino_2 + "', '" + var_codigo + "',' ', 1,0,0,0,0,0,0," + Str(rs!floa_Sal_costo) + ", " + Str(rs!FLOA_sAL_cANTIDAD) + ",0,0,0,0,0,'" + CStr(rs!INTE_sAL_AÑO) + "')", var_tabla, adOpenDynamic, adLockOptimistic
                           rs.MoveNext
                     Wend
                     Cadena = "select * from tb_entradas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                     var_copia = CopyFile(var_archivo_origen, var_archivo_destino, 1)
                     var_Archivo = var_almacen_origen + var_Archivo
                     rs.Close
                     rs.Open "select vcha_alm_correo from tb_almacenes where  vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_correo_electronico = IIf(IsNull(rs!vcha_alm_correo), "", rs!vcha_alm_correo)
                     Else
                        var_correo_electronico = ""
                     End If
                     rs.Close
                     var_tabla.Close
                     If var_correo_electronico <> "" Then
                        If MAPISession1.SessionID = 0 Then
                           MAPISession1.SignOn
                        End If
                        MAPIMessages1.SessionID = MAPISession1.SessionID
                        MAPIMessages1.Compose
                        MAPIMessages1.RecipDisplayName = var_correo_electronico
                        MAPIMessages1.RecipAddress = var_correo_electronico
                        MAPIMessages1.MsgSubject = "Nota de reempaque " + Trim(var_Archivo)
                        MAPIMessages1.MsgNoteText = "Se adjunta nota de envio"
                        MAPIMessages1.AttachmentPathName = var_ruta & "\" + var_Archivo + ".DBF"
                        MAPIMessages1.Send True
                        If MAPISession1.SessionID > 0 Then
                           MAPISession1.SignOff
                        End If
                     Else
                        MsgBox "No se a indicado una dirección de correo electrónico", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pudo cerrar el movimiento kanban", vbOKOnly, "ATENCION"
                  End If
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
               End If
            Else
               MsgBox "El movimiento no se puede imprimir ya que las existencias de los siguientes artículos exceden a la cantidad disponible en el almacen " + var_cadena_articulos
            End If
         End If
      Else
         MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existe el archivo reempaque.dbf en la carpeta " + var_ruta, vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "No es posible generar el archivo que se enviara a la planta, es posible que este en uso, salga completamente del sistema y vuelvalo a intentar", vbOKOnly, "ATENCION"
   On ERRO GoTo salir2:
   If var_tabla.State Then
      var_tabla.Close
   End If
salir2:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   lbl_cantidad_leida = Format(0, "###,###,##0.00")
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   var_ventana = 0
   lv_traspasossalidas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_clave_almacen_destino = ""
   txt_clave_almacen_origen = ""
   txt_clave_almacen_destino.Enabled = False
   txt_clave_almacen_origen.Enabled = True
   txt_clave_almacen_origen.SetFocus
   txt_nombre_almacen_origen = ""
   txt_nombre_almacen_destino = ""
End Sub

Private Sub Command1_Click()
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
   var_posible_kanban = 0
   lbl_cantidad_leida = Format(0, "###,###,##0.00")
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 2000
   Set var_tabla = CreateObject("ADODB.connection")
   rs.Open "select VCHA_PRI_RUTA_ARCHIVOS_ENVIAR from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_ARCHIVOS_ENVIAR), "", rs!VCHA_PRI_RUTA_ARCHIVOS_ENVIAR)
   End If
   rs.Close
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic
   var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   rs.Close
   var_estatus_movimiento = ""
   var_ventana = 0
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   txt_clave_almacen_destino.Enabled = False
   txt_clave_almacen_origen.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_salidas_reempaque)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If var_tipo_lista = 1 Then
            txt_clave_almacen_origen = lv_lista.selectedItem
            txt_nombre_almacen_origen = lv_lista.selectedItem.SubItems(1)
            txt_clave_almacen_origen.SetFocus
         End If
         If var_tipo_lista = 2 Then
            txt_clave_almacen_destino = lv_lista.selectedItem
            txt_nombre_almacen_destino = lv_lista.selectedItem.SubItems(1)
            txt_clave_almacen_destino.SetFocus
         End If
         frm_lista.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         frm_lista.Visible = False
      End If
      If var_tipo_lista = 2 Then
         frm_lista.Visible = False
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_traspasossalidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         var_ventana = 1
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      cnn.CommandTimeout = 360
      Dim var_cantidad_total_leida As Double
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
            If var_movimiento_bloqueado = 0 Then
               var_almacen_destino_tem = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_origen_tem = rs!VCHA_ALM_ALMACEN_ID
               var_posible = 1
               If var_tipo_permiso = 1 Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_2 = '" + var_almacen_origen_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               If var_posible = 1 Then
                  var_estatus_movimiento = rs!char_Emo_estatus
                  var_almacen_Destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  txt_clave_almacen_destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  txt_clave_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  lv_traspasossalidas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(2).Value
                  var_tipo_almacen = IIf(IsNull(rsaux!char_alm_tipo), "", rsaux!char_alm_tipo)
                  var_correo_electronico = IIf(IsNull(rsaux!vcha_alm_correo), "", rsaux!vcha_alm_correo)
                  txt_nombre_almacen_destino = rsaux!VCHA_ALM_NOMBRE
                  rsaux.Close
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_origen = rsaux(2).Value
                  txt_nombre_almacen_origen = rsaux(3).Value
                  rsaux.Close
                  rsaux.Open "select * from tb_temporal_salidas with (nolock) where inte_sal_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_cantidad_total_leida = 0
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_traspasossalidas.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!FLOA_sAL_cANTIDAD), 0, rsaux!FLOA_sAL_cANTIDAD)
                           rsaux2.Close
                           var_cantidad_total_leida = var_cantidad_total_leida + IIf(IsNull(rsaux!FLOA_sAL_cANTIDAD), 0, rsaux!FLOA_sAL_cANTIDAD)
                           rsaux.MoveNext:
                        End If
                     Wend
                     lbl_cantidad_leida = Format(var_cantidad_total_leida, "###,###,##0.00")
                  End If
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.Close
                  If Me.lv_traspasossalidas.ListItems.Count > 13 Then
                     Me.lv_traspasossalidas.ColumnHeaders(2).Width = 4685.22
                  Else
                     Me.lv_traspasossalidas.ColumnHeaders(2).Width = 4885.22
                  End If
                  
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
               MsgBox "El movimiento esta siendo usado por otro usuario", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      var_ventana = 0
      frm_busqueda.Visible = False
   End If
   If KeyAscii = 27 Then
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
      If var_posible_kanban = 1 Then
         If IsNumeric(Me.txt_cantidad_eliminar) Then
            Set TB_CANCELAR_RES_FUERA_DE_KANBAN = New TB_CANCELAR_RES_FUERA_DE_KANBAN
            var_inserta = TB_CANCELAR_RES_FUERA_DE_KANBAN.Anadir(Me.txt_clave_almacen_origen, var_clave_movimiento, var_numero_folio, Me.lv_traspasossalidas.selectedItem, CDbl(Me.txt_cantidad_eliminar), "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_exito = "S" Then
               var_posible = True
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               var_posible = False
            End If
         Else
            Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
            var_kanban = Me.txt_codigo
            var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_cantidad_eliminar, "", "", "", "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_es_un_kanban = "S" Then
               If Me.lv_traspasossalidas.selectedItem = var_kanban_articulo_id Then
                  Set TB_CANCELAR_RESERVACION_KANBAN = New TB_CANCELAR_RESERVACION_KANBAN
                  var_kanban = Me.txt_codigo
                  var_inserta = TB_CANCELAR_RESERVACION_KANBAN.Anadir(Me.txt_clave_almacen_origen, var_clave_movimiento, var_numero_folio, Me.txt_cantidad_eliminar, "", "")
                  var_kanban_es_un_kanban = var_kanban_es_un_kanban
                  var_kanban_almacen_id = var_kanban_almacen_id
                  var_kanban_articulo_id = var_kanban_articulo_id
                  var_kanban_exito = var_kanban_exito
                  var_kanban_mensaje = var_kanban_mensaje
                  If var_kanban_exito = "S" Then
                     var_posible = True
                  Else
                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                     frmmensaje.Show 1
                     var_posible = False
                  End If
               Else
                  frmmensaje.lbl_mensaje = "El codigo de kanban no corresponde al del artículo seleccionado"
                  frmmensaje.Show 1
                  var_posible = False
               End If
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               var_posible = False
            End If
         End If
      Else
         var_posible = True
      End If
         
      If var_posible = True Then
         If var_posible_kanban = 1 Then
            If Not IsNumeric(txt_cantidad_eliminar) Then
               Me.txt_cantidad_eliminar = 1
            End If
         End If
''' fin kanban
         If IsNumeric(txt_cantidad_eliminar) Then
            Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
            Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
            var_cantidad_eliminar = Val(txt_cantidad_eliminar)
            If var_cantidad_eliminar <= Me.lv_traspasossalidas.selectedItem.SubItems(2) * 1 Then
               Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = " + var_almacen_Destino + "and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_traspasossalidas.selectedItem + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               var_inserta = False
               var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, lv_traspasossalidas.selectedItem, 0 - Val(txt_cantidad_eliminar))
               var_inserta = False
               var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, lv_traspasossalidas.selectedItem, 0 - Val(txt_cantidad_eliminar), 2005)
               rs.Close
               lv_traspasossalidas.selectedItem.SubItems(2) = lv_traspasossalidas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
               var_ventana = 0
               var_renglon = lv_traspasossalidas.selectedItem.Index
               var_cantidad_total_leida = CDbl(lbl_cantidad_leida) - var_cantidad_eliminar
               lbl_cantidad_leida = Format(var_cantidad_total_leida, "###,###,##0.00")

               Call ilumina_grid
               frm_eliminar.Visible = False
               txt_codigo.SetFocus
            Else
               MsgBox "No se puede eliminar esta cantidad", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
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

Private Sub txt_clave_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes Destino"
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

Private Sub txt_clave_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_almacen_destino) <> "" Then
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id =  '" + txt_clave_almacen_destino + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
            txt_nombre_almacen_destino = rs!VCHA_ALM_NOMBRE
            var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
            txt_clave_almacen_destino.Enabled = False
            txt_codigo.Enabled = True
            txt_codigo.SetFocus
         Else
            MsgBox "El almacen no existe o no esta autorizado", vbOKOnly, "ATENCION"
            txt_nombre_almacen_destino = ""
            txt_clave_almacen_destino = ""
            var_almacen_Destino = ""
            txt_codigo.Enabled = False
         End If
         rs.Close
      Else
         rs.Open "select * from VW_MOVIMIENTOS_ALMACENES where vcha_alm_almacen_id =  '" + txt_clave_almacen_destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'  order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         If Not rs.EOF Then
            txt_nombre_almacen_destino = rs!VCHA_ALM_NOMBRE
            var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
            txt_clave_almacen_destino.Enabled = False
            txt_codigo.Enabled = True
            txt_codigo.SetFocus
         Else
            MsgBox "El almacen no existe o no esta autorizado", vbOKOnly, "ATENCION"
            txt_nombre_almacen_destino = ""
            txt_clave_almacen_destino = ""
            var_almacen_Destino = ""
            txt_codigo.Enabled = False
         End If
         rs.Close
      End If
   End If
            
   End If
End Sub

Private Sub txt_clave_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
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
End Sub

Private Sub txt_clave_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_clave_almacen_origen) <> "" Then
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id =  '" + txt_clave_almacen_origen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
               txt_nombre_almacen_origen = rs!VCHA_ALM_NOMBRE
               var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
               Me.txt_clave_almacen_origen.Enabled = False
               txt_clave_almacen_destino.Enabled = True
               txt_clave_almacen_destino.SetFocus
            Else
               MsgBox "El almacen no existe o no esta autorizado", vbOKOnly, "ATENCION"
               txt_nombre_almacen_origen = ""
               txt_clave_almacen_origen = ""
               var_almacen_origen = ""
               txt_clave_almacen_destino.Enabled = False
            End If
            rs.Close
         Else
            rs.Open "select * from vw_movimientos_almacenes where vcha_alm_almacen_id =  '" + txt_clave_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
            If Not rs.EOF Then
               txt_nombre_almacen_origen = rs!VCHA_ALM_NOMBRE
               var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
               txt_clave_almacen_destino.Enabled = True
               txt_clave_almacen_origen.Enabled = False
               txt_clave_almacen_destino.SetFocus
            Else
               MsgBox "El almacen no existe o no esta autorizado", vbOKOnly, "ATENCION"
               txt_nombre_almacen_origen = ""
               txt_clave_almacen_origen = ""
               var_almacen_origen = ""
               txt_clave_almacen_destino.Enabled = False
            End If
            rs.Close
         End If
      End If
   End If
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
   Dim var_recontable As Integer
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If var_posible_kanban = 1 Then
         Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
         var_kanban = Me.txt_codigo
         var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_codigo, "", "", "", "", "")
         var_kanban_es_un_kanban = var_kanban_es_un_kanban
         var_kanban_almacen_id = var_kanban_almacen_id
         var_kanban_articulo_id = var_kanban_articulo_id
         var_kanban_exito = var_kanban_exito
         var_kanban_mensaje = var_kanban_mensaje
         
         If var_kanban_es_un_kanban = "S" Then
            Me.txt_codigo = var_kanban_articulo_id
         Else
            var_kanban_almacen_id = Me.txt_clave_almacen_origen
         End If
         If var_kanban_almacen_id = Me.txt_clave_almacen_origen Then
            If var_empresa = 16 Then
               If Len(Me.txt_codigo) = 6 Then
                  Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-"
               Else
                  If Len(Me.txt_codigo) = 7 Then
                     Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-" + Mid(Me.txt_codigo, 7, 1)
                  End If
               End If
            End If
            
            var_verificador = True
            If Len(Trim(txt_codigo)) = 12 Then
               Call calcula_verificador(Trim(txt_codigo))
            End If
            If var_verificador = True Then
               var_es_caja = False
               If Trim(txt_codigo) <> "" Then
                  If Left(Trim(txt_codigo), 1) = "C" Then
                     x = Mid(txt_codigo, 2, 6)
                     var_embarque_caja = 0
                     If IsNumeric(x) Then
                        var_embarque_caja = CDbl(x)
                        If var_embarque_caja = var_numero_embarque Then
                           var_es_caja = True
                        Else
                           frmmensaje.lbl_mensaje = "La caja pertenece a otro embarque"
                           frmmensaje.Show 1
                           'MsgBox "La caja pertenece al embarque " + CStr(var_embarque_caja)
                           var_es_caja = False
                        End If
                     Else
                        frmmensaje.lbl_mensaje = "Caja incorrecta"
                        frmmensaje.Show 1
                        'MsgBox "Caja incorrecta", vbOKOnly, "ATENCION"
                        var_es_caja = False
                     End If
                  Else
                     var_es_caja = False
                  End If
                  If var_es_caja = True Then
                     txt_foco.Enabled = True
                     txt_foco.SetFocus
                  Else
                     var_caja = Left(txt_codigo, 6)
                     If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000010" Or var_caja = "000011" Or var_caja = "000012" Or var_caja = "000013" Or var_caja = "000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000020" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                        var_cantidad_caja = CInt(var_caja)
                        txt_codigo = Mid(txt_codigo, 7, 5)
                     End If
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_descripcion_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
                        If IsNull(rs(43).Value) Then
                           var_recontable = 0
                        Else
                           var_recontable = rs(43).Value
                        End If
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
                        rs.Close
                        rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           txt_codigo = rs(0).Value
                           rs.Close
                           rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_descripcion_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
                              If var_cantidad_caja = 0 Then
                                 If IsNull(rs(43).Value) Then
                                    var_recontable = 0
                                 Else
                                    var_recontable = rs(43).Value
                                 End If
                              Else
                                 var_recontable = 0
                              End If
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
                              frmmensaje.lbl_mensaje = "El artículo no existe"
                              frmmensaje.Show 1
                              'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                              txt_codigo = ""
                           End If
                        Else
                           frmmensaje.lbl_mensaje = "El artículo no existe"
                           frmmensaje.Show 1
                          'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                           txt_codigo = ""
                           rs.Close
                        End If
                     End If
                  End If
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Error en Código"
               frmmensaje.Show 1
               ' MsgBox "Error en Código", vbOKOnly, "ATENCION"
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El almacén del Kanban no pertenece al almacén del movimiento"
            frmmensaje.Show 1
         End If
      Else
''' FIN KANBAN
      
      
        
         var_verificador = True
         If Len(Trim(txt_codigo)) = 12 Then
            Call calcula_verificador(Trim(txt_codigo))
         End If
         If var_verificador = True Then
            var_caja = Left(txt_codigo, 6)
            If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
               var_cantidad_caja = CInt(var_caja)
               txt_codigo = Mid(txt_codigo, 7, 5)
            End If
            var_costo = 0
            var_precio = 0
            If Trim(txt_codigo) <> "" Then
               rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If IsNull(rs(43).Value) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs(43).Value
                  End If
                  var_descripcion_articulo = rs(1).Value
                  var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                  var_precio = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
                  rs.Close
                  rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_costo = IIf(IsNull(rs(4).Value), 0, rs(4).Value)
                  End If
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
                  rs.Close
                  rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     txt_codigo = rs(0).Value
                     rs.Close
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        If var_cantidad_caja = 0 Then
                           If IsNull(rs(43).Value) Then
                              var_recontable = 0
                           Else
                              var_recontable = rs(43).Value
                           End If
                        Else
                           var_recontable = 0
                        End If
                        var_descripcion_articulo = rs(1).Value
                        var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                        var_precio = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
                        rs.Close
                        rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_costo = IIf(IsNull(rs(4).Value), 0, rs(4).Value)
                        End If
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
                        frmmensaje.lbl_mensaje = "El artículo no existe"
                        frmmensaje.Show 1
                        'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                     End If
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show 1
                     'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                     rs.Close
                  End If
               End If
            Else
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "Error en Código"
            frmmensaje.Show 1
            'MsgBox "Error en Código", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      var_pase_existencias = 1
      If var_empresa = "18" Or var_empresa = "31" Then
         If var_numero_folio = 0 Or Trim(Me.txt_folio) = "" Then
            var_cantidad_temporal = 0
         Else
            rsaux.Open "select isnull(floa_sal_cantidad,0) from tb_Temporal_salidas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_cantidad_temporal = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
            Else
               var_cantidad_temporal = 0
            End If
            rsaux.Close
         End If
         'MsgBox CStr(var_cantidad_temporal)
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select floa_exi_Cantidad_disponible from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_cantidad_Existencias = IIf(IsNull(rsaux!floa_Exi_Cantidad_disponible), 0, rsaux!floa_Exi_Cantidad_disponible)
         Else
            var_cantidad_Existencias = 0
         End If
         rsaux.Close
         var_cantidad_posible = var_cantidad_Existencias - (var_cantidad_temporal + var_cantidad_leida)
         If var_cantidad_posible < 0 Then
            var_pase_existencias = 0
         End If
      End If
      If var_empresa = "18" Then
         If Me.txt_codigo = "360010000002" Or Me.txt_codigo = "360020000009" Or Me.txt_codigo = "900000000003" Or Me.txt_codigo = "911110000005" Then
            var_pase_existencias = True
         End If
      End If
      If var_pase_existencias = 1 Then
         bandera_suma = False
         If var_primera_vez = True Then
            var_inserta = False
            var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 1)
            var_numero_folio = var_numero_folio_regreso
            txt_folio = var_numero_folio
            var_primera_vez = False
         End If
         If var_posible_kanban = 1 Then
            Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
            Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
            If var_kanban_es_un_kanban = "S" Then
               var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, Me.txt_clave_almacen_origen, Me.txt_codigo, "", "")
               If var_kanban_exito = "S" Then
                  var_posible_leido = 1
               Else
                  var_posible_leido = 0
               End If
            Else
               var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, Me.txt_clave_almacen_origen, Me.txt_codigo, "", "")
               If var_kanban_exito = "S" Then
                  var_posible_leido = 1
               Else
                  var_posible_leido = 0
               End If
            End If
         Else
            var_kanban_mensaje = ""
            var_posible_leido = 1
         End If
         If var_posible_leido = 1 Then
      
      
            Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_inserta = False
               var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, 2005)
               var_inserta = False
               var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
               rs.Close
               var_cantidad_total_leida = CDbl(lbl_cantidad_leida) + var_cantidad_leida
               lbl_cantidad_leida = Format(var_cantidad_total_leida, "###,###,##0.00")
               valor = Trim(txt_codigo)
               Set itmfound = lv_traspasossalidas.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
               lv_traspasossalidas.selectedItem.SubItems(2) = lv_traspasossalidas.selectedItem.SubItems(2) + var_cantidad_leida
               var_renglon = lv_traspasossalidas.selectedItem.Index
               Call ilumina_grid
            Else
               var_inserta = False
               var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", var_almacen_origen, 2005)
               var_inserta = False
               var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", 0, 0)
               rs.Close
               var_cantidad_total_leida = CDbl(lbl_cantidad_leida) + var_cantidad_leida
               lbl_cantidad_leida = Format(var_cantidad_total_leida, "###,###,##0.00")
               Set list_item = lv_traspasossalidas.ListItems.Add(, , Trim(txt_codigo))
               list_item.SubItems(1) = var_descripcion_articulo
               list_item.SubItems(2) = var_cantidad_leida
               var_renglon = lv_traspasossalidas.ListItems.Count
               Call ilumina_grid
            End If
            If Me.lv_traspasossalidas.ListItems.Count > 13 Then
               Me.lv_traspasossalidas.ColumnHeaders(2).Width = 4685.22
            Else
               Me.lv_traspasossalidas.ColumnHeaders(2).Width = 4885.22
            End If
         Else
            frmmensaje.lbl_mensaje = var_kanban_mensaje
            frmmensaje.Show 1
            txt_codigo = ""
         End If
      Else
         Me.txt_codigo = ""
         frmmensaje.lbl_mensaje = "La cantidad excede a la cantidad en existencias"
         frmmensaje.Show 1
      End If
      txt_codigo.SetFocus
   End If
End Sub



Private Sub txt_nombre_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes Destino"
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

Private Sub txt_nombre_almacen_destino_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_nombre_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
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
End Sub

Private Sub txt_nombre_almacen_origen_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
