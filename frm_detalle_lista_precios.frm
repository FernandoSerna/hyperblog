VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdetalle_lista_precios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de lista de precios"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   7545
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1410
      TabIndex        =   29
      Top             =   390
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   30
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   31
         Top             =   120
         Width           =   5610
      End
   End
   Begin MSComctlLib.StatusBar barra 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   28
      Top             =   7080
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_lista 
      Height          =   315
      Left            =   8010
      TabIndex        =   22
      Top             =   720
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   105
      TabIndex        =   18
      Top             =   1995
      Width           =   7320
      Begin VB.Frame Frame5 
         Height          =   60
         Left            =   15
         TabIndex        =   27
         Top             =   450
         Width           =   7290
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frm_detalle_lista_precios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Picture         =   "frm_detalle_lista_precios.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   135
         Width           =   330
      End
      Begin VB.TextBox txt_porcentaje 
         Height          =   315
         Left            =   2310
         TabIndex        =   24
         Top             =   900
         Width           =   1950
      End
      Begin VB.ComboBox cmb_listas_de_precios 
         Height          =   315
         ItemData        =   "frm_detalle_lista_precios.frx":0294
         Left            =   3345
         List            =   "frm_detalle_lista_precios.frx":0296
         TabIndex        =   20
         Top             =   555
         Width           =   3840
      End
      Begin VB.TextBox txt_clave_lista_base 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   555
         Width           =   1020
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agregar Porcentaje al Precio:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   930
         Width           =   2070
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Importar precios desde Lista:"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   21
         Top             =   615
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   105
      TabIndex        =   9
      Top             =   405
      Width           =   7320
      Begin VB.TextBox txt_nombre_catalogo 
         Height          =   315
         Left            =   2595
         MaxLength       =   50
         TabIndex        =   14
         Top             =   870
         Width           =   4605
      End
      Begin VB.TextBox txt_fecha 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1200
         Width           =   1350
      End
      Begin VB.TextBox txt_catalogo 
         Height          =   315
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   13
         Top             =   870
         Width           =   1350
      End
      Begin VB.TextBox txt_precio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   12
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   10
         Top             =   210
         Width           =   1350
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   210
         Width           =   4605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Muerte:"
         Height          =   195
         Left            =   135
         TabIndex        =   35
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Catálogo:"
         Height          =   195
         Left            =   165
         TabIndex        =   34
         Top             =   915
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   585
         Width           =   495
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frm_detalle_lista_precios.frx":0298
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frm_detalle_lista_precios.frx":039A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frm_detalle_lista_precios.frx":049C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frm_detalle_lista_precios.frx":056E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Picture         =   "frm_detalle_lista_precios.frx":0670
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7050
      Picture         =   "frm_detalle_lista_precios.frx":0772
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   3675
      Left            =   105
      TabIndex        =   0
      Top             =   3315
      Width           =   7335
      Begin VB.Frame frm_busqueda 
         Height          =   780
         Left            =   555
         TabIndex        =   32
         Top             =   1635
         Width           =   2700
         Begin VB.TextBox txt_busqueda 
            Height          =   405
            Left            =   105
            TabIndex        =   33
            Top             =   225
            Width           =   2460
         End
      End
      Begin MSComctlLib.ListView lv_detalle_lista_precios 
         Height          =   3465
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción "
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Precio                  "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Catalogo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   " nombre catalogo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "fecha muerte"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   75
      TabIndex        =   8
      Top             =   285
      Width           =   7365
   End
End
Attribute VB_Name = "frmdetalle_lista_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_Cambio As Integer
Dim var_tipo_movimiento As Integer
Dim var_tipo_lista As Integer

Private Sub cmb_listas_de_precios_Click()
   If cmb_listas_de_precios = "IMPORTAR DESDE CATALOGO DE ARTICULOS" Then
      txt_clave_lista_base = "00"
   Else
      txt_clave_lista_base = Obtener_llave(cnn, rs, "TB_listadeprecios", "VCHA_LIS_NOMBRE", cmb_listas_de_precios, 0, "T")
   End If
End Sub

Private Sub cmb_listas_de_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_porcentaje.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmd_aceptar_Click()
Dim numero_items_articulos As Variant
Dim var_contador_articulos As Integer
Dim var_signo As String
Dim var_porcentaje_1 As String
Dim var_porcentaje_2 As Double
Dim var_precio_base As Double
Dim var_precio As Double

      si = MsgBox("¿Esta seguro de importar la lista con el porcentaje correspondiente?", vbYesNo, "ATENCION")
      If si = 6 Then
         Set TB_DETALLE_LISTA_PRECIOS = New TB_DETALLE_LISTA_PRECIOS
         var_signo = Mid(txt_porcentaje, 1, 1)
         var_porcentaje_1 = Trim(Mid(txt_porcentaje, 2, 15))
         If IsNumeric(var_porcentaje_1) Then
            var_porcentaje_2 = CDbl(var_porcentaje_1)
            Dim list_item As ListItem
            If Trim(txt_porcentaje) = "" Then
            End If
            If Trim(txt_clave_lista_base) <> "" Then
               If Trim(txt_clave_lista_base) = "00" Then
                  rs.Open "select count(*)from  tb_articulos", cnn, adOpenDynamic, adLockOptimistic
                  var_contador_articulos = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
                  rs.Close
                  rs.Open "select * from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
                  numero_items_articulos = 0
                  While Not rs.EOF
                     var_precio_base = IIf(IsNull(rs!mone_art_precio_base), 0, rs!mone_art_precio_base)
                     If var_signo = "+" Then
                        var_precio = var_precio_base + (var_precio_base * (var_porcentaje_2 / 100))
                     End If
                     If var_signo = "-" Then
                        var_precio = var_precio_base - (var_precio_base * (var_porcentaje_2 / 100))
                     End If
                     barra.Panels.Item(1).Text = "Cargando Catálogo de Artículos " + Str(var_porcentaje - 1) + " %. Favor de Esperar."
                     var_anadir = TB_DETALLE_LISTA_PRECIOS.Anadir(txt_lista, rs!vcha_Art_articulo_id, var_precio)
                     Set list_item = lv_detalle_lista_precios.ListItems.Add(, , rs!vcha_Art_articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
                     list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                     numero_items_articulos = numero_items_articulos + 1
                     var_porcentaje = Round((numero_items_articulos * 100) / var_contador_articulos)
                     rs.MoveNext
                     frmdetalle_lista_precios.Refresh
                  Wend
                  rs.Close
                  MsgBox "Se a terminado el proceso de importación", vbOKOnly, "ATENCIO"
                  barra.Panels.Item(1).Text = "Se a terminado el proceso de importación"
               Else
                  rs.Open "select count(*) from VW_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + txt_clave_lista_base + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_contador_articulos = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
                  rs.Close
                  rs.Open "select * from VW_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + txt_clave_lista_base + "'", cnn, adOpenDynamic, adLockOptimistic
                  numero_items_articulos = 0
                  While Not rs.EOF
                     var_precio_base = IIf(IsNull(rs!floa_dli_precio), 0, rs!floa_dli_precio)
                     If var_signo = "+" Then
                        var_precio = var_precio_base + (var_precio_base * (var_porcentaje_2 / 100))
                     End If
                     If var_signo = "-" Then
                        var_precio = var_precio_base - (var_precio_base * (var_porcentaje_2 / 100))
                     End If
                     var_anadir = TB_DETALLE_LISTA_PRECIOS.Anadir(txt_lista, rs!vcha_Art_articulo_id, var_precio)
                     Set list_item = lv_detalle_lista_precios.ListItems.Add(, , rs!vcha_Art_articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
                     list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                     numero_items_articulos = numero_items_articulos + 1
                     var_porcentaje = Round((numero_items_articulos * 100) / var_contador_articulos)
                     rs.MoveNext
                     frmdetalle_lista_precios.Refresh
                  Wend
                  rs.Close
                  MsgBox "Se a terminado el proceso de importación", vbOKOnly, "ATENCIO"
               End If
            Else
               MsgBox "No se a seleccionado una lista base", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Porcentaje Invalido", vbOKOnly, "ATENCION"
         End If
      End If
End Sub

Private Sub cmd_deshacer_Click()
   var_tipo_movimiento = 0
   txt_codigo.Enabled = False
   txt_descripcion.Enabled = False
End Sub

Private Sub cmd_eliminar_Click()
   Set TB_DETALLE_LISTA_PRECIOS = New TB_DETALLE_LISTA_PRECIOS
   If Trim(Me.txt_codigo) <> "" Then
      si = MsgBox("¿Deseas eliminar el registro", vbYesNo, "ATENCION")
      If si = 6 Then
         var_eliminar = TB_DETALLE_LISTA_PRECIOS.Eliminar(txt_lista, txt_codigo)
         lv_detalle_lista_precios.ListItems.Remove (lv_detalle_lista_precios.selectedItem.Index)
      End If
      lv_detalle_lista_precios.SetFocus
      var_tipo_movimiento = 0
      txt_codigo.Enabled = False
      txt_descripcion.Enabled = False
   Else
      MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
   Set TB_DETALLE_LISTA_PRECIOS = New TB_DETALLE_LISTA_PRECIOS
   Dim var_modificar As Boolean
   Dim var_insertar As Boolean
   Dim var_existe As Boolean
   Dim var_posible As Boolean
   If Trim(txt_precio) = "" Then
      txt_precio = 0
   End If
   If var_tipo_movimiento = 0 Then
      var_modificar = TB_DETALLE_LISTA_PRECIOS.Modificar(txt_lista, txt_codigo, txt_precio)
      If var_modificar = True Then
         lv_detalle_lista_precios.selectedItem.SubItems(2) = Format(txt_precio, "###,###,##0.00")
               If IsDate(Me.txt_fecha) Then
                  var_dia = CStr(Day(CDate(txt_fecha)))
                  var_mes = CStr(Month(CDate(txt_fecha)))
                  var_año = CStr(Year(CDate(txt_fecha)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  rs.Open "update tb_detalle_lista_precios set vcha_cat_catalogo_id = '" + txt_catalogo + "', dtim_lis_fecha_muerte = " + var_fecha + " where vcha_lis_lista_precios_id = '" + txt_lista + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               Else
                  rs.Open "update tb_detalle_lista_precios set vcha_cat_catalogo_id = '" + txt_catalogo + "' where vcha_lis_lista_precios_id = '" + txt_lista + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
         
         MsgBox "Informacion Guardada Correctamente!", vbOKOnly + vbInformation, "Aviso"
      End If
   End If
   If var_tipo_movimiento = 1 Then
      rs.Open "select * from tb_articulos where vcha_art_articulo_id ='" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
      Else
         var_posible = False
      End If
      rs.Close
      rs.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id ='" + txt_codigo + "' and vcha_lis_lista_precios_id = '" + txt_lista + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_existe = False
      Else
         var_existe = True
      End If
      rs.Close
      If var_posible = True Then
         If var_existe = True Then
            var_insertar = TB_DETALLE_LISTA_PRECIOS.Anadir(txt_lista, txt_codigo, txt_precio)
            If var_insertar = True Then
               If IsDate(Me.txt_fecha) Then
                  var_dia = CStr(Day(CDate(txt_fecha)))
                  var_mes = CStr(Month(CDate(txt_fecha)))
                  var_año = CStr(Year(CDate(txt_fecha)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  rs.Open "update tb_detalle_lista_precios set vcha_cat_catalogo_id = '" + txt_catalogo + "', dtim_lis_fecha_muerte = " + var_fecha + " where vcha_lis_lista_precios_id = '" + txt_lista + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               Else
                  rs.Open "update tb_detalle_lista_precios set vcha_cat_catalogo_id = '" + txt_catalogo + "' where vcha_lis_lista_precios_id = '" + txt_lista + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               Set list_item = lv_detalle_lista_precios.ListItems.Add(, , txt_codigo)
               list_item.SubItems(1) = txt_descripcion
               list_item.SubItems(2) = txt_precio
               list_item.SubItems(3) = Me.txt_catalogo
               list_item.SubItems(4) = Me.txt_nombre_catalogo
               list_item.SubItems(5) = Me.txt_fecha
            End If
            n = lv_detalle_lista_precios.ListItems.Count
            lv_detalle_lista_precios.ListItems.Item(n).Selected = True
            lv_detalle_lista_precios.SetFocus
         Else
            MsgBox "El artículo ya existe en la lista", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
      End If
   End If
   var_tipo_movimiento = 0
   txt_codigo.Enabled = False
   txt_descripcion.Enabled = False
End Sub

Private Sub cmd_imprimir_Click()
            Set reporte = appl.OpenReport(App.Path + "\REP_lista_precios.rpt")
            reporte.RecordSelectionFormula = "{VW_LISTA_PRECIOS.VCHA_LIS_LISTA_ID} = '" + Me.txt_lista + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Lista de Precios"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_lista_precios.rpt")
               reporte.RecordSelectionFormula = "{VW_LISTA_PRECIOS.VCHA_LIS_LISTA_ID} = '" + Me.txt_lista + "'"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_lista_precios_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_codigo = ""
   txt_descripcion = ""
   txt_precio = ""
   txt_codigo.Enabled = True
   var_tipo_movimiento = 1
   txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_tipo_lista = 0 Then
         Me.frm_lista.Visible = False
         Me.frm_busqueda.Visible = False
      End If
   End If
End Sub

Private Sub Form_Load()
   var_tipo_lista = 0
   frm_busqueda.Visible = False
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 2300
   Dim si As Integer
   txt_codigo.Enabled = False
   txt_descripcion.Enabled = False
   var_tipo_movimiento = 0
   rs.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_clave_lista + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      rs.Close
      rs.Open "select count(*) from VW_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_clave_lista + "'", cnn, adOpenDynamic, adLockOptimistic
      var_contador_articulos = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
      rs.Close
      rs.Open "select * from VW_detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_clave_lista + "'", cnn, adOpenDynamic, adLockOptimistic
      numero_items_articulos = 0
      If Not rs.EOF Then
         While Not rs.EOF
            Set list_item = lv_detalle_lista_precios.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_dli_precio), 0, rs!floa_dli_precio), "###,###,##0.00")
            list_item.SubItems(3) = IIf(IsNull(rs!vcha_cat_catalogo_id), "", rs!vcha_cat_catalogo_id)
            list_item.SubItems(4) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            list_item.SubItems(5) = IIf(IsNull(rs!dtim_lis_fecha_muerte), "", rs!dtim_lis_fecha_muerte)
            numero_items_articulos = numero_items_articulos + 1
            var_porcentaje = Round((numero_items_articulos * 100) / var_contador_articulos)
            rs.MoveNext
            frmdetalle_lista_precios.Refresh
         Wend
      Else
         MsgBox "La fecha de la lista de precios ya caduco", vbOKOnly, "ATENCION"
         'MsgBox cnn.ConnectionString
         rsaux4.Open "select * from VW_DETALLE_LISTA_PRECIOS_CADUCA where VCHA_LIS_LISTA_PRECIOS_ID = '" + var_clave_lista + "'", cnn, adOpenDynamic, adLockOptimistic
         numero_items_articulos = 0
         If Not rsaux4.EOF Then
            While Not rsaux4.EOF
                  Set list_item = lv_detalle_lista_precios.ListItems.Add(, , rsaux4!vcha_Art_articulo_id)
                  list_item.SubItems(1) = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
                  list_item.SubItems(2) = Format(IIf(IsNull(rsaux4!floa_dli_precio), 0, rsaux4!floa_dli_precio), "###,###,##0.00")
                  list_item.SubItems(3) = IIf(IsNull(rs!vcha_cat_catalogo_id), "", rs!vcha_cat_catalogo_id)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
                  list_item.SubItems(5) = IIf(IsNull(rs!dtim_lis_fecha_muerte), "", rs!dtim_lis_fecha_muerte)
                  numero_items_articulos = numero_items_articulos + 1
                  'var_porcentaje = Round((numero_items_articulos * 100) / var_contador_articulos)
                  rsaux4.MoveNext
                  frmdetalle_lista_precios.Refresh
             Wend
          End If
          rsaux4.Close
      End If
      rs.Close
      
      txt_clave_lista_base.Enabled = False
      cmb_listas_de_precios.Enabled = False
      cmd_aceptar.Enabled = False
      cmd_cancelar.Enabled = False
      If lv_detalle_lista_precios.ListItems.Count > 0 Then
         txt_codigo = lv_detalle_lista_precios.selectedItem
         txt_descripcion = lv_detalle_lista_precios.selectedItem.SubItems(1)
         txt_precio = lv_detalle_lista_precios.selectedItem.SubItems(2)
         txt_catalogo = lv_detalle_lista_precios.selectedItem.SubItems(3)
         Me.txt_nombre_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(4)
         Me.txt_fecha = Me.lv_detalle_lista_precios.selectedItem.SubItems(5)
      End If
      var_tipo_movimiento = 0
   Else
      rs.Close
      si = MsgBox("¿Deseas importar una lista de precios?", vbYesNo, "ATENCION")
      If si = 6 Then
         txt_clave_lista_base.Enabled = True
         cmb_listas_de_precios.Enabled = True
         cmd_aceptar.Enabled = True
         cmd_cancelar.Enabled = True
         rsaux.Open "select * from tb_listadeprecios", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_listas_de_precios.hwnd, rsaux, 1)
         rsaux.Close
         cmb_listas_de_precios.AddItem ("IMPORTAR DESDE CATALOGO DE ARTICULOS")
      Else
      End If
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_detalle_lista_precios)
End Sub

Private Sub lv_detalle_lista_precios_Click()
   txt_codigo = lv_detalle_lista_precios.selectedItem
   txt_descripcion = lv_detalle_lista_precios.selectedItem.SubItems(1)
   txt_precio = lv_detalle_lista_precios.selectedItem.SubItems(2)
   Me.txt_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(3)
   Me.txt_nombre_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(4)
   Me.txt_fecha = Me.lv_detalle_lista_precios.selectedItem.SubItems(5)
   var_tipo_movimiento = 0
End Sub

Private Sub lv_detalle_lista_precios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_detalle_lista_precios, ColumnHeader)
End Sub

Private Sub lv_detalle_lista_precios_GotFocus()
   On Error GoTo salir:
   txt_codigo = lv_detalle_lista_precios.selectedItem
   txt_descripcion = lv_detalle_lista_precios.selectedItem.SubItems(1)
   txt_precio = lv_detalle_lista_precios.selectedItem.SubItems(2)
   Me.txt_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(3)
   Me.txt_nombre_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(4)
   Me.txt_fecha = Me.lv_detalle_lista_precios.selectedItem.SubItems(5)
   var_tipo_movimiento = 0
salir:
End Sub

Private Sub lv_detalle_lista_precios_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_codigo = lv_detalle_lista_precios.selectedItem
   txt_descripcion = lv_detalle_lista_precios.selectedItem.SubItems(1)
   txt_precio = lv_detalle_lista_precios.selectedItem.SubItems(2)
   Me.txt_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(3)
   Me.txt_nombre_catalogo = Me.lv_detalle_lista_precios.selectedItem.SubItems(4)
   Me.txt_fecha = Me.lv_detalle_lista_precios.selectedItem.SubItems(5)
   var_tipo_movimiento = 0
End Sub

Private Sub opt_cambiar_formula_Click()
   var_tipo_Cambio = 2
   txt_nuevo_precio = ""
   txt_nuevo_precio.Enabled = False
   txt_nuevo_formula = ""
   txt_nuevo_formula.Enabled = True
   txt_nuevo_formula.SetFocus
End Sub

Private Sub opt_cambiar_precio_Click()
  var_tipo_Cambio = 1
  txt_nuevo_formula = ""
  txt_nuevo_formula.Enabled = False
  txt_nuevo_precio = ""
  txt_nuevo_precio.Enabled = True
  txt_nuevo_precio.SetFocus
End Sub

Private Sub lv_detalle_lista_precios_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      txt_busqueda = ""
      frm_busqueda.Visible = True
      txt_busqueda.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_codigo = lv_lista.selectedItem
            txt_descripcion = lv_lista.selectedItem.SubItems(1)
            rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + lv_lista.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_precio = IIf(IsNull(rs!mone_art_precio_base), 0, rs!mone_art_precio_base)
            Else
               Me.txt_precio = 0
            End If
            rs.Close
         End If
         txt_codigo.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            Me.txt_catalogo = Me.lv_lista.selectedItem
            Me.txt_nombre_catalogo = Me.lv_lista.selectedItem.SubItems(1)
            Me.txt_catalogo.SetFocus
         End If
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
   var_tipo_lista = 0
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_busqueda) <> "" Then
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_busqueda + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux2.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_clave_lista + "' AND VCHA_ART_ARTICULO_ID = '" + txt_busqueda + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               valor = txt_busqueda
               Set itmfound = Me.lv_detalle_lista_precios.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
            Else
               MsgBox "El articulo no esta asignado a esta lista de precios", vbOKOnly, "ATENCION"
            End If
            rsaux2.Close
         Else
            rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_busqueda + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               txt_busqueda = rsaux!vcha_Art_articulo_id
               rsaux2.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_clave_lista + "' AND VCHA_ART_ARTICULO_ID = '" + txt_busqueda + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  valor = txt_busqueda
                  Set itmfound = Me.lv_detalle_lista_precios.findItem(valor, lvwText, , lvwPartial)
                  itmfound.EnsureVisible
                  itmfound.Selected = True
               Else
                  MsgBox "El articulo no esta asignado a esta lista de precios", vbOKOnly, "ATENCION"
               End If
               rsaux2.Close
            Else
               MsgBox "El articulo no existe", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         End If
         rs.Close
      End If
      frm_busqueda.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_LostFocus()
   frm_busqueda.Visible = False
End Sub

Private Sub txt_catalogo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
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

Private Sub txt_catalogo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_lista_base_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clave_lista_base_LostFocus()
   If Trim(txt_clave_lista_base) <> "" Then
      rs.Open "select * from tb_listadeprecios where vcha_lis_lista_precio_id = '" + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_listas_de_precios = rs!vcha_lis_nobre
         txt_clave_lista_base.Enabled = False
         cmb_listas_de_precios.Enabled = False
      Else
         txt_procentaje.Enabled = True
      End If
      rs.Close
   End If
End Sub


Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900
      Else
         lv_lista.ColumnHeaders(2).Width = 4100
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   Dim var_posible As Boolean
   If Trim(txt_codigo) <> "" Then
      var_posible = False
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
         txt_descripcion = rs!vcha_art_nombre_español
         txt_precio = rs!mone_art_precio_base
         rs.Close
         rs.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + txt_lista + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_precio = IIf(IsNull(rs!floa_dli_precio), 0, rs!floa_dli_precio)
         End If
         rs.Close
      Else
         rs.Close
         rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_posible = True
               txt_codigo = rs!vcha_Art_articulo_id
               txt_descripcion = rsaux!vcha_art_nombre_español
               txt_precio = rsaux!mone_art_precio_base
               rsaux.Close
               rs.Close
               rs.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + txt_lista + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_precio = IIf(IsNull(rs!floa_dli_precio), 0, rs!floa_dli_precio)
               End If
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
      If var_posible = True Then
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         txt_codigo.SetFocus
      End If
   End If
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900
      Else
         lv_lista.ColumnHeaders(2).Width = 4100
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_catalogo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
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

Private Sub txt_nombre_catalogo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 45, 43
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 45
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub
