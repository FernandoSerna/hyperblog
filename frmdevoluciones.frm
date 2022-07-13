VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmdevoluciones 
   Caption         =   "devoluciones"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   Icon            =   "frmdevoluciones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1335
      Index           =   0
      Left            =   4260
      TabIndex        =   36
      Top             =   1095
      Width           =   3210
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
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   420
         Width           =   1500
      End
      Begin VB.TextBox txt_folio_enviado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   37
         Top             =   945
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   41
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número Movimiento:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   40
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número Enviado:"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   39
         Top             =   990
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   35
      Top             =   960
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Index           =   1
      Left            =   -180
      TabIndex        =   29
      Top             =   2310
      Width           =   4140
      Begin VB.TextBox txt_almacen_origen 
         Height          =   345
         Left            =   780
         TabIndex        =   31
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txt_almacen_destino 
         Height          =   345
         Left            =   780
         TabIndex        =   30
         Top             =   855
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   34
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   33
         Top             =   510
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   32
         Top             =   120
         Width           =   4065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4860
      Left            =   45
      TabIndex        =   17
      Top             =   2430
      Width           =   7425
      Begin VB.TextBox txt_cantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         TabIndex        =   22
         Top             =   495
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   19
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   20
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
            TabIndex        =   21
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
         TabIndex        =   18
         Top             =   465
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_traspasosentradas 
         Height          =   3750
         Left            =   45
         TabIndex        =   23
         Top             =   1035
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   6615
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6085
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Enviaron"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Recibidos"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "COSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Diferencia"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   615
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   570
      Width           =   7455
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   7845
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3615
      Width           =   1125
   End
   Begin VB.Frame frm_numero_traspaso 
      Height          =   1185
      Left            =   -660
      TabIndex        =   9
      Top             =   -90
      Width           =   4455
      Begin VB.TextBox txt_numero_traspaso 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   780
         Width           =   2760
      End
      Begin VB.ComboBox cmb_movimientos_salidas 
         Height          =   315
         Left            =   1305
         TabIndex        =   10
         Top             =   435
         Width           =   3075
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Número de traspaso"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   45
         TabIndex        =   14
         Top             =   120
         Width           =   4380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número de traspaso:"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   825
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mov. de Salida:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   465
         Width           =   1110
      End
   End
   Begin VB.Frame frm_leer 
      Height          =   1245
      Left            =   345
      TabIndex        =   3
      Top             =   1095
      Width           =   4560
      Begin VB.TextBox txt_leer 
         Height          =   300
         Left            =   810
         TabIndex        =   5
         Top             =   450
         Width           =   2295
      End
      Begin VB.ComboBox cmb_almacen_destino 
         Height          =   315
         Left            =   795
         TabIndex        =   4
         Top             =   810
         Width           =   3600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Archivo:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   8
         Top             =   465
         Width           =   585
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de archivos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   45
         TabIndex        =   7
         Top             =   120
         Width           =   4470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   6
         Top             =   810
         Width           =   585
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   330
      TabIndex        =   0
      Top             =   1095
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   465
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
         Width           =   3075
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Index           =   1
      Left            =   7080
      TabIndex        =   27
      Top             =   690
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   420
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
            Picture         =   "frmdevoluciones.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":201A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":28F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":31D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":3AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":3BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":3CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":3DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":3EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdevoluciones.frx":4004
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Index           =   0
      Left            =   90
      TabIndex        =   28
      Top             =   690
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Movimiento"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Movimiento"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Movimiento"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Leer archivo"
            ImageIndex      =   2
         EndProperty
      EndProperty
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
      TabIndex        =   42
      Top             =   75
      Width           =   7350
   End
End
Attribute VB_Name = "frmdevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen_destino As String
Dim var_almacen_origen As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Integer
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_movimiento_salida As String
Dim var_numero_salida As Integer
Dim var_tabla As ADODB.Connection
Dim var_ruta As String

Private Sub cmb_almacen_destino_Click()
   var_almacen_destino = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 2, "T")
   var_tipo_almacen = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 10, "T")
End Sub

Private Sub cmb_almacen_destino_KeyPress(KeyAscii As Integer)
   Set TB_TRASPASOS_INSERTA = New TB_TRASPASOS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   If KeyAscii = 13 Then
      If Trim(cmb_almacen_destino) <> "" Then
         rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero_origen = " + Str(var_numero_salida) + " and vcha_emo_almacen_origen = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            MsgBox "La nota numero " + txt_leer + " ya fue leida en el movimiento " + Str(rs(3).Value), vbOKOnly, "ATENCION"
            rs.Close
            txt_folio_enviado = ""
         Else
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = 'XD' and inte_emo_numero_origen = " + Str(var_numero_salida) + " and vcha_emo_almacen_origen = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_inserta_traspaso = 1
            Else
               var_inserta_traspaso = 0
            End If
            rs.Close
            If var_inserta_traspaso = 0 Then
               rs.Open "select max(INTE_EMO_NUMERO) as numero from tb_encabezado_movimientos where VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               If IsNull(rs(0).Value) Then
                  var_numero_folio = 1
                  var_inserta = False
                  var_insreta = TB_ENCABEZADO_MOVIMIENTOS_I.Anadir(var_almacen_origen, "XD", Now, var_numero_salida, Str(var_numero_salida), "", "", "", var_almacen_origen, var_almacen_destino, "I", fun_NombreUsuario, fun_NombrePc, "", "XD", "")
               Else
                  var_numero_folio = rs(0).Value + 1
                  var_inserta = False
                  var_insreta = TB_ENCABEZADO_MOVIMIENTOS_I.Anadir(var_almacen_origen, "XD", Now, var_numero_salida, Str(var_numero_salida), "", "", "", var_almacen_origen, var_almacen_destino, "I", fun_NombreUsuario, fun_NombrePc, "", "XD", "")
               End If
               rs.Close
               txt_folio = var_numero_folio
               rs.Open "select cvetienda,folio,codigo,cant1,costo from " + txt_leer, var_tabla, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs(2).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_precio = rsaux(2).Value
                  Else
                     var_precio = 0
                  End If
                  var_inserta = False
                  var_inserta = TB_SALIDAS_INSERTA.Anadir(var_almacen_origen, "XD", var_numero_salida, rs(2).Value, rs(3).Value, rs(4).Value, var_precio, 0)
                  var_inserta = False
                  var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_almacen_origen, "XD", var_numero_salida, rs(2).Value, rs(3).Value, rs(4).Value, var_precio, 0)
                  var_inserta = False
                  var_inserta = TB_TRASPASOS_INSERTA.Anadir(var_almacen_origen, "XD", var_numero_salida, rs(2).Value, rs(3).Value, 0, rs(4).Value, var_precio, 0, var_almacen_origen)
                  rsaux.Close
                  rs.MoveNext
               Wend
               rs.Close
            End If
            var_primera_vez = True
            rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_almacen_destino = rsaux(2).Value
            txt_almacen_destino.Text = rsaux(3).Value
            txt_almacen_destino.Enabled = False
            rsaux.Close
            rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_almacen_origen = rsaux(2).Value
            txt_almacen_origen.Text = rsaux(3).Value
            txt_almacen_origen.Enabled = False
            rsaux.Close
            var_movimiento_salida = "XD"
            rsaux.Open "select * from tb_TRASPASOS where inte_tra_numero = " + Str(var_numero_salida) + " and vcha_mov_movimiento_id = '" + "XD" + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               lv_traspasosentradas.ListItems.Clear
               While Not rsaux.EOF
                  rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux(3).Value + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     Set list_item = lv_traspasosentradas.ListItems.Add(, , rsaux(3).Value)
                     list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                     list_item.SubItems(2) = IIf(IsNull(rsaux(4).Value), "", rsaux(4).Value)
                     list_item.SubItems(3) = IIf(IsNull(rsaux(5).Value), "", rsaux(5).Value)
                     list_item.SubItems(4) = IIf(IsNull(rsaux(6).Value), "", rsaux(6).Value)
                     list_item.SubItems(5) = list_item.SubItems(2) - list_item.SubItems(3)
                  End If
                  rsaux2.Close
                  rsaux.MoveNext:
               Wend
               txt_codigo.Enabled = True
               txt_codigo.SetFocus
            End If
            rsaux.Close
            var_primera_vez = True
            frm_leer.Visible = False
         End If
      End If
    Else
       MsgBox "No se a seleccionado ningun almacen para su destino", vbOKOnly, "ATENCION"
    End If
    If KeyAscii = 27 Then
      frm_leer.Visible = False
   End If
   KeyAscii = 0

End Sub

Private Sub cmb_almacen_destino_LostFocus()
   var_tabla.Close
   frm_leer.Visible = False
End Sub

Private Sub cmb_movimientos_salidas_Click()
   var_movimiento_salida = Obtener_llave(cnn, rsaux, "TB_MOVIMIENTOS", "VCHA_MOV_NOMBRE", cmb_movimientos_salidas, 0, "T")
End Sub

Private Sub cmb_movimientos_salidas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_numero_traspaso.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_numero_traspaso.Visible = False
   End If
End Sub

Private Sub Form_Load()
   var_estatus_movimiento = ""
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   frm_numero_traspaso.Visible = False
   frm_leer.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   txt_almacen_origen = ""
   txt_almacen_destino = ""
   txt_almacen_origen.Enabled = False
   txt_almacen_destino.Enabled = False
End Sub

Private Sub lv_traspasosentradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub lv_traspasosentradas_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set var_tabla = CreateObject("ADODB.connection")
   If Index = 0 Then
      Select Case Button.Index
         Case 1
            txt_codigo.Enabled = False
            var_primera_vez = True
            frm_busqueda.Visible = False
            lv_traspasosentradas.ListItems.Clear
            var_numero_folio = 0
            txt_folio = ""
            txt_codigo = ""
            var_estatus_movimiento = ""
            var_movimiento_salida = ""
            If var_tipo_permiso = 1 Then
               rs.Open "select * from vw_movimientos_permisos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and char_mov_afectacion = 'TS'", cnn, adOpenDynamic, adLockOptimistic
               Call RecsetToCombo(cmb_movimientos_salidas.hwnd, rs, 1)
               rs.Close
            Else
               rs.Open "select * from tb_movimientos  where char_mov_afectacion = 'TS' order by VCHA_mov_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
               Call RecsetToCombo(cmb_movimientos_salidas.hwnd, rs, 1)
               rs.Close
            End If
            frm_numero_traspaso.Visible = True
            cmb_movimientos_salidas.SetFocus
         Case 2
            frm_busqueda.Visible = True
            txt_busqueda_folio.SetFocus
         Case 3
            If var_numero_folio > 0 Then
               If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENTRADAS_TRASPASOS.VCHA_EMO_MOVIMIENTO_ORIGEN} = '" + var_movimiento_salida + "' and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO_ORIGEN} = " + Str(var_numero_salida) + " and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
                  frmvistasprevias.cr.ReportSource = reporte
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
               Else
                  var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
                  If var_si = 1 Then
                     Cadena = "select * from tb_TRASPASOS where vcha_alm_almacen_id = " + var_almacen_origen + " and  VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' and inte_TRA_numero = " + Str(var_numero_salida)
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                         var_inserta = False
                         var_inserta = TB_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, rs(5).Value, rs(7).Value, rs(8).Value, rs(9).Value, rs(10).Value, var_almacen_origen)
                         rs.MoveNext
                     Wend
                     rs.Close
                     var_estatus_movimiento = "I"
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", "", var_almacen_destino, "I", fun_NombreUsuario, fun_NombrePc, Now)
                     Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos.rpt")
                     reporte.RecordSelectionFormula = "{VW_ENTRADAS_TRASPASOS.VCHA_EMO_MOVIMIENTO_ORIGEN} = '" + var_movimiento_salida + "' and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO_ORIGEN} = " + Str(var_numero_salida) + " and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
                     frmvistasprevias.cr.ReportSource = reporte
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     txt_codigo.Enabled = False
                     txt_foco.Enabled = False
                  End If
               End If
            Else
               MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
            End If
         Case 4
            cmb_almacen_destino.Enabled = False
            frm_leer.Visible = True
            txt_leer.SetFocus
      End Select
   End If
   If Index = 1 Then
      Unload Me
      frmcodigo_acceso.Show
   End If
ErrHandler:
   
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
       frm_busqueda.Visible = False
   End If
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_almacen_destino_tem = rs!VCHA_EMO_ALMACEN_DESTINO
            var_almacen_origen_tem = rs!VCHA_EMO_ALMACEN_ORIGEN
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
               var_numero_folio = rs!inte_emo_numero
               var_numero_salida = rs!INTE_EMO_NUMERO_ORIGEN
               var_movimiento_salida = rs!VCHA_EMO_MOVIMIENTO_ORIGEN
               var_almacen_destino = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_origen = rs!VCHA_EMO_ALMACEN_ORIGEN
               var_estatus_movimiento = rs!CHAR_EMO_ESTATUS
               rs.Close
               var_primera_vez = False
               lv_traspasosentradas.ListItems.Clear
               txt_folio_enviado = var_numero_salida
               txt_folio = var_numero_folio
               rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_almacen_destino = rsaux(2).Value
               txt_almacen_destino.Text = rsaux(3).Value
               txt_almacen_destino.Enabled = False
               rsaux.Close
               rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_almacen_origen = rsaux(2).Value
               txt_almacen_origen.Text = rsaux(3).Value
               txt_almacen_origen.Enabled = False
               rsaux.Close
               rsaux.Open "select * from tb_TRASPASOS where inte_TRA_numero = " + Str(var_numero_salida) + " and vcha_mov_movimiento_id = '" + var_movimiento_salida + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                     rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Set list_item = lv_traspasosentradas.ListItems.Add(, , rsaux!vcha_art_articulo_id)
                        list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                        list_item.SubItems(2) = IIf(IsNull(rsaux!floa_tra_cantidad), 0, rsaux!floa_tra_cantidad)
                        list_item.SubItems(3) = IIf(IsNull(rsaux!floa_tra_cantidad_recibida), 0, rsaux!floa_tra_cantidad_recibida)
                        list_item.SubItems(4) = IIf(IsNull(rsaux!floa_tra_costo), 0, rsaux!floa_tra_costo)
                        list_item.SubItems(5) = list_item.SubItems(2) - list_item.SubItems(3)
                        rsaux2.Close
                        rsaux.MoveNext:
                     End If
                  Wend
                  frm_busqueda.Visible = False
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     txt_codigo.Enabled = False
                  Else
                     txt_codigo.Enabled = True
                     txt_codigo.SetFocus
                  End If
               Else
                  MsgBox "El traspaso no a sido enviado o no a sido impreso", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            Else
               MsgBox "No esta autorizado para afectar este movimiento"
               rs.Close
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
            rs.Close
         End If
      End If
      frm_numero_traspaso.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Set TB_TRASPASOS_MODIFICA = New TB_TRASPASOS_MODIFICA
      var_cantidad_eliminar = Val(txt_cantidad_eliminar)
      var_inserta = False
      var_inserta = TB_TRASPASOS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_movimiento_salida, var_numero_salida, lv_traspasosentradas.SelectedItem, 0 - Val(txt_cantidad_eliminar), var_almacen_origen)
      lv_traspasosentradas.SelectedItem.SubItems(3) = lv_traspasosentradas.SelectedItem.SubItems(3) - Val(txt_cantidad_eliminar)
      lv_traspasosentradas.SelectedItem.SubItems(5) = lv_traspasosentradas.SelectedItem.SubItems(5) + Val(txt_cantidad_eliminar)
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = 1#
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
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

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
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
            var_costo = rs(3).Value
            var_precio = rs(2).Value
            rs.Close
            rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_costo = rs(4).Value
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
                  If IsNull(rs(43).Value) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs(43).Value
                  End If
                  var_descripcion_articulo = rs(1).Value
                  var_costo = rs(3).Value
                  var_precio = rs(2).Value
                  rs.Close
                  rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_costo = rs(4).Value
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
                  MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               rs.Close
            End If
         End If
      Else
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Set TB_TRASPASOS_INSERTA = New TB_TRASPASOS_INSERTA
   Set TB_TRASPASOS_MODIFICA = New TB_TRASPASOS_MODIFICA
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      bandera_suma = False
      If var_primera_vez = True Then
         rs.Open "select max(INTE_EMO_NUMERO) as numero from tb_encabezado_movimientos where VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If IsNull(rs(0).Value) Then
            var_numero_folio = 1
            var_inserta = False
            var_insreta = TB_ENCABEZADO_MOVIMIENTOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, Now, var_numero_folio, Str(var_numero_salida), "", "", "", var_almacen_origen, var_almacen_destino, "", fun_NombreUsuario, fun_NombrePc, "", var_movimiento_salida, "")
         Else
            var_numero_folio = rs(0).Value + 1
            var_inserta = False
            var_insreta = TB_ENCABEZADO_MOVIMIENTOS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, Now, var_numero_folio, Str(var_numero_salida), "", "", "", var_almacen_origen, var_almacen_destino, "", fun_NombreUsuario, fun_NombrePc, "", var_movimiento_salida, "")
         End If
         rs.Close
         txt_folio = var_numero_folio
         var_primera_vez = False
      End If
      Cadena = "select * from TB_TRASPASOS where vcha_alm_almacen_id = " + var_almacen_origen + "and  VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' and inte_tra_numero = " + Str(var_numero_salida) + " and vcha_art_articulo_id = '" + txt_codigo + "' AND VCHA_TRA_ALMACEN_ORIGEN = '" + var_almacen_origen + "'"
      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_inserta = False
         var_costo = lv_traspasosentradas.SelectedItem.SubItems(4)
         var_inserta = TB_TRASPASOS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_movimiento_salida, var_numero_salida, txt_codigo, var_cantidad_leida, var_almacen_origen)
         rs.Close
         valor = Trim(txt_codigo)
         Set itmfound = lv_traspasosentradas.FindItem(valor, lvwText, , lvwPartial)
         itmfound.EnsureVisible
         itmfound.Selected = True
         lv_traspasosentradas.SelectedItem.SubItems(3) = lv_traspasosentradas.SelectedItem.SubItems(3) + var_cantidad_leida
         lv_traspasosentradas.SelectedItem.SubItems(5) = lv_traspasosentradas.SelectedItem.SubItems(2) - lv_traspasosentradas.SelectedItem.SubItems(3)
         
      Else
         var_inserta = False
         var_inserta = TB_TRASPASOS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_movimiento_salida, var_numero_salida, txt_codigo, 0, var_cantidad_leida, var_costo, var_precio, "0", var_almacen_origen)
         rs.Close
         Set list_item = lv_traspasosentradas.ListItems.Add(, , Trim(txt_codigo))
         list_item.SubItems(1) = var_descripcion_articulo
         list_item.SubItems(2) = 0
         list_item.SubItems(3) = var_cantidad_leida
         list_item.SubItems(4) = var_costo
         list_item.SubItems(5) = list_item.SubItems(2) - list_item.SubItems(3)
      End If
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_leer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error GoTo ersalir:
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_almacen_destino.hwnd, rs, 3)
         rs.Close
      Else
         rs.Open "select * from tb_almacenes order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_almacen_destino.hwnd, rs, 3)
         rs.Close
      End If
      rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockOptimistic
      var_ruta = rs!VCHA_PRI_RUTA_NOTAS_ENVIO
      rs.Close
      var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Tablas de Visual FoxPro;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
      rs.Open "select cvetienda,folio,codigo,cant1,costo from " + txt_leer, var_tabla, adOpenDynamic, adLockOptimistic
      var_almacen_origen_tem = rs(0).Value
      var_posible = 1
      If var_tipo_permiso = 1 Then
      
      End If
      If var_posible = 1 Then
         var_almacen_origen = rs(0).Value
         var_numero_salida = rs(1).Value
         txt_folio_enviado = var_numero_salida
         cmb_almacen_destino.Enabled = True
         cmb_almacen_destino.SetFocus
         rs.Close
      Else
         rs.Close
         MsgBox "No esta autorizado para leer archivos de este almacen", vbOKOnly, "ATENCION"
         frm_leer.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_leer.Visible = False
   End If
   Exit Sub
ersalir:
   MsgBox "A surgido un error al leer el archivo, puede que el archivo este siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
   frm_leer.Visible = False
End Sub

Private Sub txt_numero_traspaso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_numero_traspaso.Visible = False
   End If
   If KeyAscii = 13 Then
      If Trim(txt_numero_traspaso) <> "" Then
         rs.Open "select * from tb_encabezado_movimientos where VCHA_EMO_MOVIMIENTO_ORIGEN = '" + var_movimiento_salida + "' and INTE_EMO_NUMERO_ORIGEN = " + txt_numero_traspaso, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            MsgBox "El traspaso ya fue cargado en el movimiento número " + Str(rs(3).Value), vbOKOnly, "ATENCION"
            rs.Close
         Else
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_numero_traspaso + " and vcha_mov_movimiento_id = '" + var_movimiento_salida + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_almacen_destino_tem = rs(9).Value
               var_almacen_origen_tem = rs(8).Value
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
                  var_numero_salida = Val(txt_numero_traspaso)
                  txt_folio_enviado = var_numero_salida
                  var_almacen_destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  var_almacen_origen = rs!VCHA_EMO_ALMACEN_ORIGEN
                  lv_traspasosentradas.ListItems.Clear
                  var_primera_vez = True
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(2).Value
                  txt_almacen_destino.Text = rsaux(3).Value
                  txt_almacen_destino.Enabled = False
                  rsaux.Close
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_origen = rsaux(2).Value
                  txt_almacen_origen.Text = rsaux(3).Value
                  txt_almacen_origen.Enabled = False
                  rsaux.Close
                  rsaux.Open "select * from tb_TRASPASOS where inte_TRA_numero = " + txt_numero_traspaso + " and vcha_mov_movimiento_id = '" + var_movimiento_salida + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_traspasosentradas.ListItems.Add(, , rsaux!vcha_art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!floa_tra_cantidad), 0, rsaux!floa_tra_cantidad)
                           list_item.SubItems(3) = IIf(IsNull(rsaux!floa_tra_cantidad_recibida), 0, rsaux!floa_tra_cantidad_recibida)
                           list_item.SubItems(4) = IIf(IsNull(rsaux!floa_tra_costo), 0, rsaux!floa_tra_costo)
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                     txt_codigo.Enabled = True
                     txt_codigo.SetFocus
                  Else
                     MsgBox "El traspaso no a sido enviado o no a sido impreso", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      End If
      frm_numero_traspaso.Visible = False
   End If
End Sub

