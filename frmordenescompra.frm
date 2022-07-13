VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmordenescompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Compra"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmordenescompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7365
   Begin VB.CommandButton cmd_costos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmordenescompra.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Costos Predeterminados"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   960
      TabIndex        =   54
      Top             =   765
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   55
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
         TabIndex        =   56
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6765
      Picture         =   "frmordenescompra.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_ruta 
      Height          =   3615
      Left            =   7485
      TabIndex        =   47
      Top             =   1200
      Width           =   3330
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   1680
         TabIndex        =   53
         Top             =   3210
         Width           =   1575
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   75
         TabIndex        =   52
         Top             =   3210
         Width           =   1605
      End
      Begin VB.TextBox txt_path 
         Height          =   330
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   450
         Width           =   3180
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   75
         TabIndex        =   49
         Top             =   870
         Width           =   3180
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   90
         TabIndex        =   48
         Top             =   2820
         Width           =   3195
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Seleccione la Ruta de la Empresa"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   50
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame frm_importacion 
      Height          =   1290
      Left            =   7590
      TabIndex        =   40
      Top             =   990
      Width           =   5310
      Begin VB.TextBox txt_folio 
         Height          =   315
         Left            =   1185
         TabIndex        =   45
         Top             =   840
         Width           =   2160
      End
      Begin VB.TextBox txt_ruta_empresa 
         Height          =   315
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   495
         Width           =   3570
      End
      Begin MSComctlLib.Toolbar tlb_busqueda_ruta 
         Height          =   330
         Index           =   0
         Left            =   4800
         TabIndex        =   46
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Busqueda de Ruta"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   135
         TabIndex        =   44
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ruta empresa:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   555
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Importación de Información"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   41
         Top             =   120
         Width           =   5235
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmordenescompra.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Orden Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmordenescompra.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Orden Alt + B"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmordenescompra.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Orden Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_importar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      Picture         =   "frmordenescompra.frx":130C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Importar Información Alt + M"
      Top             =   15
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_articulos 
      Height          =   3015
      Left            =   960
      TabIndex        =   35
      Top             =   1980
      Width           =   5550
      Begin VB.ListBox lst_articulos 
         Height          =   2790
         Left            =   75
         TabIndex        =   36
         Top             =   150
         Width           =   5415
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   540
      TabIndex        =   31
      Top             =   285
      Width           =   3135
      Begin VB.TextBox txt_busqueda_numero 
         Height          =   315
         Left            =   210
         TabIndex        =   32
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Número de Orden de Compra"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   33
         Top             =   120
         Width           =   3075
      End
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Index           =   0
      Left            =   4425
      TabIndex        =   28
      Top             =   660
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   72810497
      CurrentDate     =   37581
   End
   Begin VB.TextBox txt_foco 
      Height          =   345
      Left            =   8115
      TabIndex        =   16
      Top             =   2595
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Height          =   1650
      Index           =   0
      Left            =   150
      TabIndex        =   22
      Top             =   465
      Width           =   7035
      Begin VB.TextBox txt_nombre_proveedor 
         Height          =   315
         Left            =   2415
         TabIndex        =   8
         Top             =   870
         Width           =   4425
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2415
         TabIndex        =   10
         Top             =   1215
         Width           =   4425
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   6495
         Picture         =   "frmordenescompra.frx":140E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Seleccione la fecha"
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   525
         Width           =   1065
      End
      Begin VB.TextBox txt_numero 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   525
         Width           =   2100
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1125
         TabIndex        =   9
         Top             =   1215
         Width           =   1260
      End
      Begin VB.TextBox txt_proveedor 
         Height          =   315
         Left            =   1125
         TabIndex        =   7
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   2
         Left            =   4890
         TabIndex        =   27
         Top             =   585
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   26
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Left            =   300
         TabIndex        =   25
         Top             =   1275
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   24
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Datos Generales "
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   6960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5310
      Left            =   150
      TabIndex        =   18
      Top             =   2070
      Width           =   7035
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   3225
         TabIndex        =   37
         Top             =   2565
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   39
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   38
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSMask.MaskEdBox txt_costo 
         Height          =   330
         Left            =   5700
         TabIndex        =   14
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   582
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txt_cantidad 
         Height          =   315
         Left            =   3765
         TabIndex        =   13
         Top             =   548
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_descripcion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         TabIndex        =   15
         Top             =   915
         Width           =   5760
      End
      Begin VB.TextBox txt_codigo 
         Height          =   360
         Left            =   1155
         TabIndex        =   11
         Top             =   510
         Width           =   1665
      End
      Begin MSComctlLib.ListView lv_ordenescompra 
         Height          =   3915
         Left            =   60
         TabIndex        =   17
         Top             =   1305
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   6906
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   5935
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Costo"
            Object.Width           =   1799
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   975
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Costo:"
         Height          =   195
         Index           =   2
         Left            =   5175
         TabIndex        =   30
         Top             =   608
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   1
         Left            =   3030
         TabIndex        =   29
         Top             =   608
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   608
         Width           =   540
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Detalle de Orden de Compra"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Width           =   6960
      End
   End
   Begin MSComDlg.CommonDialog cmdentradas 
      Left            =   6810
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Busqueda de archivo"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6000
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":1510
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":1DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":26C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":2C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":353C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":3E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":46F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4802
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4914
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4D5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   21
      Top             =   285
      Width           =   7065
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordenescompra.frx":4E6E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmordenescompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_numero As Integer
Dim var_primera_vez As Boolean
Dim var_descripcion_articulo As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_origen_codigo As Integer
Dim var_fecha_orden As Date
Dim var_estatus As String
Dim var_ruta As String
Dim var_tipo_lista As Integer

Private Sub cmb_almacen_Change()

End Sub

Private Sub cmb_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cmb_proveedor_Change()
   
End Sub

Private Sub cmb_proveedor_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cmb_almacenes_Click()
   txt_almacen = Obtener_llave(cnn, rs, "TB_ALMACENES", "VCHA_ALM_NOMBRE", cmb_almacenes, 2, "T")
   txt_codigo.Enabled = True
End Sub

Private Sub cmb_almacenes_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False

End Sub

Private Sub cmb_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_almacen) <> "" Then
         txt_codigo.SetFocus
      Else
         cmb_almacenes.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_proveedores_Click()
   txt_proveedor = Obtener_llave(cnn, rs, "TB_PROVEEDORES", "VCHA_PRO_NOMBRE", cmb_proveedores, 0, "T")
   cmb_almacenes.Enabled = True
End Sub

Private Sub cmb_proveedores_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub cmb_proveedores_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txt_proveedor)) > 0 Then
         txt_proveedor.Enabled = False
         cmb_proveedores.Enabled = False
         txt_almacen.Enabled = True
         txt_almacen.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmd_aceptar_Click()
   txt_ruta_empresa = txt_path
   frm_ruta.Visible = False
End Sub

Private Sub cmd_buscar_Click()
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Dim var_referencia_oc As String
   txt_almacen.Enabled = False
   txt_proveedor.Enabled = False
   frm_ruta.Visible = False
   frm_importacion.Visible = False
   txt_busqueda_numero = ""
   frm_busqueda.Visible = True
   txt_busqueda_numero.SetFocus
End Sub

Private Sub cmd_costos_Click()
   var_activa_forma_costos_predeterminados = Me.Name
   frmcostos_predeterminados.Show
   Me.Enabled = False
End Sub

Private Sub cmd_importar_Click()
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Dim var_referencia_oc As String
   frm_ruta.Visible = False
   frm_importacion.Visible = False
             frm_importacion.Visible = True
             txt_folio.SetFocus
             
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Dim var_referencia_oc As String
   frm_ruta.Visible = False
   frm_importacion.Visible = False
   If var_numero > 0 Then
      If var_estatus = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\orden_compra.rpt")
         reporte.RecordSelectionFormula = "{VW_ORDENES_COMPRA.INTE_OCO_NUMERO} = " & var_numero
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Orden de Compra"
         frmvistasprevias.Show
         Set reporte = Nothing
      Else
         rs.Open "SELECT * FROM TB_ORDENES_COMPRA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND INTE_OCO_NUMERO = " + Str(var_numero), cnn, adOpenDynamic, adLockOptimistic
         var_estatus = "I"
         If Len(Trim(Str(var_numero))) = 1 Then
            var_referencia_oc = "00000" + Trim(Str(var_numero))
         End If
         If Len(Trim(Str(var_numero))) = 2 Then
            var_referencia_oc = "0000" + Trim(Str(var_numero))
         End If
         If Len(Trim(Str(var_numero))) = 3 Then
            var_referencia_oc = "000" + Trim(Str(var_numero))
         End If
         If Len(Trim(Str(var_numero))) = 4 Then
            var_referencia_oc = "00" + Trim(Str(var_numero))
         End If
         If Len(Trim(Str(var_numero))) = 5 Then
            var_referencia_oc = "0" + Trim(Str(var_numero))
         End If
         If Len(Trim(Str(var_numero))) = 6 Then
            var_referencia_oc = Trim(Str(var_numero))
         End If
         var_referencia_oc = "EUP" + var_referencia_oc
         If Not rs.EOF Then
            'cnn.BeginTrans
            While Not rs.EOF
               ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, rs!VCHA_ALM_ALMACEN_ID, "EU", rs!inte_oco_numero, Date, "P", rs!VCHA_PRO_PROVEEDOR_ID, rs!vcha_Art_articulo_id, rs!floa_oco_costo, rs!floa_oco_cantidad, 0, "", var_referencia_oc, 0, 0, 2005, "", 0)
               rs.MoveNext
            Wend
            rsaux2.Open "update tb_ordenes_compra set char_oco_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_oco_numero = " + Str(var_numero), cnn, adOpenDynamic, adLockOptimistic
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            'cnn.CommitTrans
            Set reporte = appl.OpenReport(App.Path + "\orden_compra.rpt")
            reporte.RecordSelectionFormula = "{VW_ORDENES_COMPRA.INTE_OCO_NUMERO} = " & var_numero
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Orden de Compra"
            frmvistasprevias.Show
            Set reporte = Nothing
         Else
            MsgBox "No existe la orden de compra número " & var_numero, vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      txt_codigo.Enabled = False
      txt_cantidad.Enabled = False
      txt_foco.Enabled = False
      txt_costo.Enabled = False
   Else
      MsgBox "No se a seleccionado una orden de compra", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
   Dim var_referencia_oc As String
   var_estatus = ""
   txt_numero = ""
   txt_folio = ""
   txt_almacen = ""
   txt_nombre_almacen = ""
   txt_proveedor = ""
   txt_nombre_proveedor = ""
   frm_ruta.Visible = False
   frm_importacion.Visible = False
   txt_proveedor.Enabled = True
   txt_almacen.Enabled = True
   txt_proveedor.SetFocus
   Me.txt_fecha = Date
   var_fecha_orden = Date
   Me.lv_ordenescompra.ListItems.Clear
   var_primera_vez = True
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmdfecha_Click(Index As Integer)
   frm_ruta.Visible = False
   frm_importacion.Visible = False
   mes(0).Value = Date
   mes(0).Visible = True
End Sub

Private Sub Command1_Click()
   frm_ruta.Visible = False
End Sub

Private Sub Dir1_Change()
   txt_path = Dir1.Path
End Sub

Private Sub Dir1_Click()
   txt_path = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_ruta.Visible = False
   Else
      txt_path = Dir1.Path
   End If
End Sub

Private Sub Drive1_Change()
   On Error GoTo salir:
   Dir1.Path = Drive1.Drive
   Exit Sub
salir:
   MsgBox "La unidad " + Drive1.Drive + " no esta disponible", vbOKOnly, "ATENCION"
   Drive1.Drive = "c:"
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_ruta.Visible = False
   End If
End Sub

Private Sub Form_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 66 Then
      cmd_buscar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 77 Then
      cmd_importar_Click
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
     If Me.frm_articulos.Visible = True Then
        Me.frm_articulos.Visible = False
        If Me.txt_codigo.Enabled = True Then
           Me.txt_codigo.SetFocus
        End If
     Else
        If Me.frm_busqueda.Visible = True Then
           Me.frm_busqueda.Visible = False
        Else
           If Me.frm_eliminar.Visible = True Then
              Me.frm_eliminar.Visible = False
           Else
              If Me.mes(0).Visible = True Then
                 Me.mes(0).Visible = False
              Else
                 Unload Me
              End If
           End If
        End If
     End If
   End If
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 2000
   frm_ruta.Visible = False
   frm_importacion.Visible = False
   frm_eliminar.Visible = False
   frm_articulos.Visible = False
   txt_descripcion = ""
   txt_foco.Enabled = False
   frm_busqueda.Visible = False
   mes(0).Visible = False
   txt_fecha = Now
   txt_almacen.Enabled = False
   txt_proveedor.Enabled = False
   txt_codigo.Enabled = False
   txt_cantidad.Enabled = False
   txt_costo.Enabled = False
   var_primera_vez = True
   var_numero = 0
   txt_proveedor.Enabled = False
   txt_almacen.Enabled = False
   var_fecha_orden = Date
   var_estatus = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_ordenescompra)
End Sub

Private Sub lst_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_articulos.Visible = False
      var_oigen_codigo = 0
   End If
   If KeyAscii = 13 Then
      var_origen_codigo = 1
      txt_codigo = Obtener_llave(cnn, rs, "TB_ARTICULOS", "VCHA_ART_NOMBRE_ESPAÑOL", lst_articulos.Text, 0, "T")
      txt_descripcion = lst_articulos.Text
      frm_articulos.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_proveedor = lv_lista.selectedItem
            txt_nombre_proveedor = lv_lista.selectedItem.SubItems(1)
         Else
            txt_proveedor = ""
            txt_nombre_proveedor = ""
         End If
         If txt_proveedor.Enabled = True Then
            txt_proveedor.SetFocus
         End If
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_almacen = lv_lista.selectedItem
            txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
         Else
            txt_almacen = ""
            txt_nombre_almacen = ""
         End If
         If txt_almacen.Enabled = True Then
            txt_almacen.SetFocus
         End If
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         If txt_proveedor.Enabled = True Then
            txt_proveedor.SetFocus
         End If
      End If
      If var_tipo_lista = 2 Then
         If txt_almacen.Enabled = True Then
            txt_almacen.SetFocus
         End If
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_ordenescompra_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub lv_ordenescompra_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus = "" Then
         txt_cantidad_eliminar = 0
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      Else
         MsgBox "La orden de compra ya no puede ser modificada", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_ordenescompra_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub mes_nowClick(Index As Integer, ByVal nowClicked As Date)
   txt_fecha.Text = mes(0).Value
   mes(0).Visible = False
End Sub

Private Sub Text2_Change()
   Text1 = Text2
End Sub

Private Sub mes_DateDblClick(Index As Integer, ByVal DateDblClicked As Date)
   txt_fecha = mes(0).Value
   mes(0).Visible = False
End Sub

Private Sub mes_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes(0).Visible = False
   End If
   If KeyAscii = 13 Then
      txt_fecha = mes(0).Value
      mes(0).Visible = False
   End If
End Sub

Private Sub tlb_busqueda_ruta_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
   frm_ruta.Visible = True
   Dir1.SetFocus
End Sub


Private Sub txt_almacen_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
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

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacenes_Change()

End Sub

Private Sub txt_almacen_LostFocus()
   If Len(Trim(txt_almacen)) > 0 Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
         If Trim(txt_proveedor) <> "" Then
            txt_codigo.Enabled = True
            txt_proveedor.Enabled = False
         Else
            MsgBox "Dede de indicar un proveedor", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         txt_almacen = ""
         txt_nombre_almacen = ""
         txt_almacen.Enabled = True
         txt_nombre_almacen.Enabled = True
         txt_codigo.Enabled = False
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_numero) <> "" Then
         rs.Open "select * from tb_ordenes_compra where inte_oco_numero = " + txt_busqueda_numero + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_estatus = IIf(IsNull(rs!char_oco_estatus), "", rs!char_oco_estatus)
            var_fecha_orden = rs!dtim_oco_fecha
            txt_fecha = var_fecha_orden
            txt_descripcion = ""
            var_primera_vez = False
            lv_ordenescompra.ListItems.Clear
            While Not rs.EOF
               rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
               Set list_item = lv_ordenescompra.ListItems.Add(, , rs!vcha_Art_articulo_id)
                   list_item.SubItems(1) = Trim(IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value))
                   list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_oco_cantidad), 0, rs!floa_oco_cantidad), "###,###,##0.00")
                   list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_oco_costo), 0, rs!floa_oco_costo), "###,###,##0.00")
               End If
               rsaux.Close
               var_numero = rs!inte_oco_numero
               txt_numero = var_numero
               txt_proveedor = rs!VCHA_PRO_PROVEEDOR_ID
               rsaux2.Open "SELECT * FROM TB_PROVEEDORES WHERE VCHA_PRO_PROVEEDOR_ID = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  txt_nombre_proveedor = IIf(IsNull(rsaux2!VCHA_PRO_NOMBRE), "", rsaux2!VCHA_PRO_NOMBRE)
               Else
                  txt_nombre_proveedor = ""
               End If
               rsaux2.Close
               txt_almacen = rs!VCHA_ALM_ALMACEN_ID
               rsaux2.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  Me.txt_nombre_almacen = IIf(IsNull(rsaux2!VCHA_ALM_NOMBRE), "", rsaux2!VCHA_ALM_NOMBRE)
               Else
                  Me.txt_nombre_almacen = ""
               End If
               rsaux2.Close
               txt_fecha = rs!dtim_oco_fecha
               rs.MoveNext:
             Wend
             rs.Close
             cmb_proveedores = Obtener_llave(cnn, rs, "TB_PROVEEDORES", "VCHA_PRO_PROVEEDOR_ID", txt_proveedor, 1, "T")
             cmb_almacenes = Obtener_llave(cnn, rs, "TB_ALMACENES", "VCHA_ALM_ALMACEN_ID", txt_almacen, 3, "T")
             If var_estatus = "" Then
                txt_codigo.Enabled = True
                txt_costo.Enabled = True
                txt_cantidad.Enabled = True
                txt_foco.Enabled = True
             Else
                txt_codigo.Enabled = False
                txt_costo.Enabled = False
                txt_cantidad.Enabled = False
                txt_foco.Enabled = False
             End If
         Else
            MsgBox "El número de orden de compra no existe", vbOKOnly, "ATENCION"
            rs.Close
         End If
       End If
      frm_busqueda.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_numero_LostFocus()
   If Len(Trim(txt_almacen)) > 0 Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      Else
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         txt_almacen = ""
         txt_nombre_almacen = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 52, 8, 46
    Case 13
       If IsNumeric(txt_cantidad_eliminar) Then
          If (txt_cantidad_eliminar * 1) <= (lv_ordenescompra.selectedItem.SubItems(2) * 1) Then
             rs.Open "update tb_ordenes_compra set floa_oco_cantidad = floa_oco_cantidad - " + txt_cantidad_eliminar + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_oco_numero = " + Str(var_numero) + " and vcha_Art_articulo_id = '" + Me.lv_ordenescompra.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
             If rs.State = 1 Then
                rs.Close
             End If
             lv_ordenescompra.selectedItem.SubItems(2) = Format((lv_ordenescompra.selectedItem.SubItems(2) * 1) - (txt_cantidad_eliminar * 1), "###,###,##0.00")
             frm_eliminar.Visible = False
          Else
             MsgBox "La cantidad exede a la cantidad en la orden de compra", vbOKOnly, "ATENCION"
             txt_cantidad_eliminar.SetFocus
          End If
       Else
          MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
          txt_cantidad_eliminar.SetFocus
       End If
    Case 27
       txt_cantidad_eliminar = 0
       frm_eliminar.Visible = False
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txt_cantidad_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 52, 8, 46
    Case 13
      If Len(Trim(txt_cantidad)) = 0 Then
         txt_cantidad = 0
      End If
      If IsNumeric(txt_cantidad) Then
         txt_costo.SetFocus
      Else
         MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
         txt_cantidad.SetFocus
      End If
    Case Else
      KeyAscii = 0
    End Select
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_proveedor.Enabled = False
   Me.txt_almacen.Enabled = False
   frm_ruta.Visible = False
   frm_importacion.Visible = False
   If var_origen_codigo = 0 Then
      txt_codigo.Text = ""
   End If
   txt_cantidad.Enabled = True
   txt_costo.Enabled = True
   txt_foco.Enabled = False
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      rs.Open "select vcha_art_nombre_español from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
         lst_articulos.AddItem rs(0).Value
         rs.MoveNext
      Wend
      rs.Close
      frm_articulos.Visible = True
      var_origen_codigo = 1
      lst_articulos.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_posible As Boolean
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      var_posible = False
      If Trim(txt_codigo) <> "" Then
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + Me.txt_proveedor + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               txt_costo = IIf(IsNull(rsaux!floa_cpr_costo_predeterminado), 0, rsaux!floa_cpr_costo_predeterminado)
            End If
            rsaux.Close
            var_posible = True
            txt_cantidad.Enabled = True
            txt_costo.Enabled = True
            txt_foco.Enabled = True
            var_descripcion_articulo = rs(1).Value
            txt_descripcion = rs(1).Value
            txt_cantidad.SetFocus
            var_origen_codigo = 0
            rs.Close
         Else
            rs.Close
            rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  txt_codigo = rs!vcha_Art_articulo_id
                  rsaux2.Open "select * from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + Me.txt_proveedor + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     txt_costo = IIf(IsNull(rsaux2!floa_cpr_costo_predeterminado), 0, rsaux2!floa_cpr_costo_predeterminado)
                  End If
                  rsaux2.Close
                  var_posible = True
                  txt_cantidad.Enabled = True
                  txt_costo.Enabled = True
                  txt_foco.Enabled = True
                  txt_descripcion = rsaux!vcha_art_nombre_español
                  txt_cantidad.SetFocus
                  var_origen_codigo = 0
                  var_descripcion_articulo = rsaux!vcha_art_nombre_español
                  rsaux.Close
                  rs.Close
               Else
                  var_posible = False
                  rsaux.Close
                  rs.Close
               End If
            Else
               rs.Close
            End If
         End If
      Else
         var_posible = False
      End If
      If var_posible = False Then
         MsgBox "El artículo no existe", vbOKOnly, "Atención"
         txt_cantidad.Enabled = True
         txt_costo.Enabled = True
         txt_foco.Enabled = False
         var_origen_codigo = 0
         txt_codigo.SetFocus
      End If
   End If
End Sub

Private Sub txt_costo_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub txt_costo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 52, 8, 46
    Case 13
       If Len(Trim(txt_costo)) = 0 Then
          txt_costo = 0
       End If
       If IsNumeric(txt_costo) Then
          txt_foco.Enabled = True
          txt_foco.SetFocus
       Else
          MsgBox "Costo Incorrecto", vbOKOnly, "ATENCION"
          txt_costo.SetFocus
       End If
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txt_foco_GotFocus()
   Dim var_actualiza As Boolean
   Set TB_ORDENES_COMPRA_MODIFICA = New TB_ORDENES_COMPRA_MODIFICA
   Set TB_ORDENES_COMPRA_INSERTA = New TB_ORDENES_COMPRA_INSERTA
   If txt_codigo <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      If var_primera_vez = True Then
         rs.Open "select max(INTE_OCO_NUMERO) FROM TB_ORDENES_COMPRA", cnn, adOpenDynamic, adLockOptimistic
         If IsNull(rs(0).Value) Then
            var_numero = 1
         Else
            var_numero = rs(0).Value + 1
         End If
         txt_numero = var_numero
         rs.Close
         var_primera_vez = False
      End If
      rs.Open "select * from tb_ordenes_compra where inte_oco_numero = " + Str(var_numero) + " and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_actualiza = TB_ORDENES_COMPRA_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_numero, var_fecha_orden, txt_almacen, txt_proveedor, txt_codigo, txt_cantidad, txt_costo, 0)
      Else
         var_actualiza = TB_ORDENES_COMPRA_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_numero, var_fecha_orden, txt_almacen, txt_proveedor, txt_codigo, txt_cantidad, txt_costo, 0)
         Set list_item = lv_ordenescompra.ListItems.Add(, , txt_codigo)
         list_item.SubItems(1) = var_descripcion_articulo
         list_item.SubItems(2) = Format(0, "###,###,##0.00")
         list_item.SubItems(3) = Format(txt_costo, "###,###,##0.00")
      End If
      rs.Close
      valor = txt_codigo
      Set itmfound = lv_ordenescompra.findItem(valor, lvwText, , lvwPartial)
      itmfound.EnsureVisible
      itmfound.Selected = True
      lv_ordenescompra.selectedItem.SubItems(2) = Format(lv_ordenescompra.selectedItem.SubItems(2) + Int(txt_cantidad), "###,###,##0.00")
      lv_ordenescompra.selectedItem.SubItems(3) = Format(txt_costo, "###,###,##0.00")
      txt_codigo = ""
      txt_cantidad = ""
      txt_costo = ""
   End If
   txt_codigo.SetFocus
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_importacion.Visible = False
   End If
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
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

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_proveedores order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PRO_PROVEEDOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PROVEEDORES"
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

Private Sub txt_nombre_proveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False

End Sub

Private Sub txt_path_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_ruta.Visible = False
   End If
End Sub

Private Sub txt_proveedor_GotFocus()
   frm_ruta.Visible = False
   frm_importacion.Visible = False
End Sub

Private Sub txt_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_proveedores order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PRO_PROVEEDOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PROVEEDORES"
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

Private Sub txt_proveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_proveedor_LostFocus()
   If Len(Trim(txt_proveedor)) > 0 Then
      rs.Open "select * from tb_proveedores where vcha_pro_proveedor_id = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_proveedor = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
         txt_codigo.Enabled = False
      Else
         MsgBox "Clave de proveedor incorrecta", vbOKOnly, "ATENCIO"
         txt_proveedor.Enabled = True
         txt_nombre_proveedor.Enabled = True
         txt_proveedor = ""
         txt_nombre_proveedor = ""
      End If
      rs.Close
   Else
      txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_ruta_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_importacion.Visible = False
   End If
End Sub
