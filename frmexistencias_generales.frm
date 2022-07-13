VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmexistencias_generales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Existencias"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmexistencias_generales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Caption         =   "Seleccionar "
      Height          =   1110
      Left            =   90
      TabIndex        =   45
      Top             =   1515
      Width           =   2715
      Begin VB.OptionButton opt_seleccion_linea 
         Caption         =   "Linea"
         Height          =   270
         Left            =   210
         TabIndex        =   47
         Top             =   645
         Width           =   1140
      End
      Begin VB.OptionButton opt_seleccion_articulos 
         Caption         =   "Artículo"
         Height          =   270
         Left            =   210
         TabIndex        =   46
         Top             =   330
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Filtrar Reporte Por "
      Height          =   1110
      Left            =   2955
      TabIndex        =   42
      Top             =   1515
      Width           =   2715
      Begin VB.OptionButton opt_filtrar_ninguna 
         Caption         =   "Ninguna"
         Height          =   270
         Left            =   225
         TabIndex        =   48
         Top             =   240
         Width           =   2340
      End
      Begin VB.OptionButton opt_filtrar_negativos 
         Caption         =   "Negativos"
         Height          =   270
         Left            =   210
         TabIndex        =   44
         Top             =   780
         Width           =   1140
      End
      Begin VB.OptionButton opt_filtrar_ceros 
         Caption         =   "Existencias en cero"
         Height          =   270
         Left            =   210
         TabIndex        =   43
         Top             =   510
         Width           =   2340
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Filtrar Reporte Por "
      Height          =   1110
      Left            =   2955
      TabIndex        =   18
      Top             =   405
      Width           =   2715
      Begin VB.OptionButton opt_ambos 
         Caption         =   "Ambos"
         Height          =   240
         Left            =   165
         TabIndex        =   21
         Top             =   795
         Width           =   2325
      End
      Begin VB.OptionButton opt_solo_costo 
         Caption         =   "Solo Costo"
         Height          =   240
         Left            =   165
         TabIndex        =   20
         Top             =   495
         Width           =   2325
      End
      Begin VB.OptionButton opt_solo_precio 
         Caption         =   "Solo Precio"
         Height          =   240
         Left            =   165
         TabIndex        =   19
         Top             =   195
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar Reporte Agrupado Por "
      Height          =   1110
      Left            =   90
      TabIndex        =   15
      Top             =   405
      Width           =   2715
      Begin VB.OptionButton opt_linea 
         Caption         =   "Linea"
         Height          =   270
         Left            =   210
         TabIndex        =   17
         Top             =   555
         Width           =   1140
      End
      Begin VB.OptionButton opt_general 
         Caption         =   "Artículo"
         Height          =   270
         Left            =   210
         TabIndex        =   16
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1395
      TabIndex        =   11
      Top             =   4830
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   75300865
      CurrentDate     =   37761
   End
   Begin VB.Frame Frame4 
      Caption         =   " Fecha "
      Height          =   720
      Left            =   90
      TabIndex        =   12
      Top             =   6570
      Width           =   5595
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3105
         Picture         =   "frmexistencias_generales.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11235
      Picture         =   "frmexistencias_generales.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   330
      Width           =   11640
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmexistencias_generales.frx":19AE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_canales 
      Height          =   3945
      Left            =   90
      TabIndex        =   0
      Top             =   2610
      Width           =   5595
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmexistencias_generales.frx":1AB0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   30
         Picture         =   "frmexistencias_generales.frx":1CC6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmexistencias_generales.frx":1DC8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmexistencias_generales.frx":1E9A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmexistencias_generales.frx":20E4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   3180
         Left            =   45
         TabIndex        =   6
         Top             =   690
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   5609
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
            Text            =   "Clave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Almacenes"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Frame frm_Articulo 
      Caption         =   " Artículos "
      Height          =   6885
      Left            =   5730
      TabIndex        =   22
      Top             =   405
      Width           =   5865
      Begin VB.OptionButton opt_seleccion 
         Caption         =   "Selección de artículos"
         Height          =   345
         Left            =   435
         TabIndex        =   50
         Top             =   240
         Width           =   1980
      End
      Begin VB.OptionButton opt_todos 
         Caption         =   "Todos los artículos"
         Height          =   225
         Left            =   3360
         TabIndex        =   49
         Top             =   255
         Width           =   1845
      End
      Begin VB.Frame Frame8 
         Height          =   120
         Left            =   15
         TabIndex        =   30
         Top             =   555
         Width           =   5820
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmexistencias_generales.frx":22FA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmexistencias_generales.frx":2510
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Marcar (Enter)"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmexistencias_generales.frx":275A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   105
         Picture         =   "frmexistencias_generales.frx":282C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmexistencias_generales.frx":292E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1785
         TabIndex        =   23
         Top             =   750
         Width           =   1740
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   5700
         Left            =   60
         TabIndex        =   29
         Top             =   1110
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   10054
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de artículo:"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   31
         Top             =   810
         Width           =   1575
      End
   End
   Begin VB.Frame frm_linea 
      Caption         =   " Lineas "
      Height          =   6885
      Left            =   5730
      TabIndex        =   32
      Top             =   405
      Width           =   5865
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1815
         TabIndex        =   39
         Top             =   780
         Width           =   1620
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmexistencias_generales.frx":2B44
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   105
         Picture         =   "frmexistencias_generales.frx":2D5A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmexistencias_generales.frx":2E5C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmexistencias_generales.frx":2F2E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmexistencias_generales.frx":3178
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   15
         TabIndex        =   33
         Top             =   525
         Width           =   5820
      End
      Begin MSComctlLib.ListView lv_lineas 
         Height          =   5625
         Left            =   60
         TabIndex        =   40
         Top             =   1170
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   9922
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de linea:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   41
         Top             =   795
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmexistencias_generales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_todos_articulos As Integer
Dim var_todos_lineas As Integer

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim si As Integer
   Dim var_existen_almacenes As Integer
   Dim var_cadena_linea As String
   If IsDate(txt_fecha) Then
      si = MsgBox("Se calcularan las existencias al dia " + Format(CDate(txt_fecha), "Long Date"), vbYesNo, "ATENCION")
      If si = 6 Then
         var_n = lv_almacenes.ListItems.Count
         var_existen_almacenes = 0
         For var_i = 1 To var_n
             lv_almacenes.ListItems.item(var_i).Selected = True
             If lv_almacenes.selectedItem.SubItems(2) = "*" Then
                var_existen_almacenes = var_existen_almacenes + 1
             End If
         Next var_i
         If var_existen_almacenes > 0 Then
            cnn.BeginTrans
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select max(inte_exi_consecutivo) from TB_TEMP_EXISTENCIAS_ALMACENES", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            'MsgBox "Insert into  (INTE_EXI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')"
            rs.Open "Insert into TB_TEMP_EXISTENCIAS_ENTRADAS  (INTE_EXI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            cnn.CommandTimeout = 360
            If Me.opt_seleccion_articulos = True Then
               If Me.opt_todos = True Then
                  var_todos_articulos = 1
               Else
                  var_todos_articulos = 0
               End If
               If var_todos_articulos = 1 Then
                  rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
               Else
                  var_n = lv_articulos.ListItems.Count
                  For var_i = 1 To var_n
                      lv_articulos.ListItems.item(var_i).Selected = True
                      rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) values (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                  Next var_i
               End If
            End If
            
            If Me.opt_seleccion_linea = True Then
               If var_todos_lineas = 1 Then
                  If var_empresa = "18" Then
                     rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from VW_EXISTEN_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
                  End If
               Else
                  var_n = lv_lineas.ListItems.Count
                  For var_i = 1 To var_n
                      lv_lineas.ListItems.item(var_i).Selected = True
                       If Trim(lv_lineas.selectedItem.SubItems(2)) = "*" Then
                         rsaux.Open "SELECT VCHA_ART_ARTICULO_ID from tb_articulos where vcha_lin_linea_id = '" + lv_lineas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         While Not rsaux.EOF
                               rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) values (" + CStr(var_consecutivo) + ", '" + rsaux!VCHA_ART_ARTICULO_ID + "')", cnn, adOpenDynamic, adLockOptimistic
                               rsaux.MoveNext
                         Wend
                         rsaux.Close
                      End If
                  Next var_i
               End If
            End If
            
            var_fecha_fin_1 = CDate(txt_fecha) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            
             var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_año = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             var_fecha_fin_1 = CDate(txt_fecha)
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_año = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             'var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            
            var_n = lv_almacenes.ListItems.Count
            For var_i = 1 To var_n
                lv_almacenes.ListItems.item(var_i).Selected = True
                If lv_almacenes.selectedItem.SubItems(2) = "*" Then
                   rs.Open "Insert into TB_TEMP_EXISTENCIAS_ALMACENES (INTE_EXI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina, VCHA_ALM_ALMACEN_ID) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '" + lv_almacenes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            If rsaux5.State Then
               rsaux5.Close
            End If
            
            rsaux5.Open "select * from TB_TEMP_EXISTENCIAS_ALMACENES where inte_exi_consecutivo = " + CStr(var_consecutivo) + " and vcha_alm_almacen_ID is not null", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommandTimeout = 3000
            var_filtro = ""
            While Not rsaux5.EOF
                  var_almacen = rsaux5!VCHA_ALM_ALMACEN_ID
                  If opt_general = True Then
                     If var_filtro = "" Then
                        var_filtro = var_filtro + "({VW_EXISTENCIAS_GENERALES.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     Else
                        var_filtro = var_filtro + " or {VW_EXISTENCIAS_GENERALES.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     End If
                  Else
                     If var_filtro = "" Then
                        var_filtro = var_filtro + "({VW_EXISTENCIAS_GENERALES_LINEA.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     Else
                        var_filtro = var_filtro + " or {VW_EXISTENCIAS_GENERALES_LINEA.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     End If
                  End If
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_1 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_2 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_3 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_4 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_5 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            var_filtro = var_filtro + ")"
            ' se quitaron el dia 26-09-2008
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_1 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_2 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_3 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_4 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_5 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "exec SP_EXISTENCIAS_RAPIDAS_6 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
           
                         
                         
            If opt_general = True Then
               If opt_solo_precio = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_sin_costo.rpt ")
               End If
               If opt_solo_costo = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_sin_precio.rpt")
               End If
               If opt_ambos = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales.rpt")
               End If
               VAR_CADENA_FILTRO = "{VW_EXISTENCIAS_GENERALES.INTE_EXI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_EXISTENCIAS_GENERALES.VCHA_AUD_USUARIO}= '" + var_clave_usuario_global + "' and {VW_EXISTENCIAS_GENERALES.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
               If Me.opt_filtrar_negativos = True Then
                  If opt_solo_precio = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and {VW_EXISTENCIAS_GENERALES.FLOA_EXI_CANTIDAD} >= 0"
                  End If
                  If opt_solo_costo = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and {VW_EXISTENCIAS_GENERALES.FLOA_EXI_CANTIDAD} >= 0"
                  End If
                  If opt_ambos = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and {VW_EXISTENCIAS_GENERALES.FLOA_EXI_CANTIDAD} >= 0"
                  End If
               End If
               If Me.opt_filtrar_ceros = True Then
                  If opt_solo_precio = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and {VW_EXISTENCIAS_GENERALES.FLOA_EXI_CANTIDAD} <> 0"
                  End If
                  If opt_solo_costo = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and {VW_EXISTENCIAS_GENERALES.FLOA_EXI_CANTIDAD} <> 0"
                  End If
                  If opt_ambos = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and {VW_EXISTENCIAS_GENERALES.FLOA_EXI_CANTIDAD} <> 0"
                  End If
               End If
               VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and " + var_filtro
               reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de existencias"
               frmvistasprevias.Show 1
               Set reporte = Nothing
            
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  If opt_solo_precio = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_sin_costo_excel.rpt ")
                  End If
                  If opt_solo_costo = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_sin_precio_Excel.rpt")
                  End If
                  If opt_ambos = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_Excel.rpt")
                  End If
                  VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and " + var_filtro
                  reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_existencias" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            End If
            If opt_linea = True Then
               If opt_solo_precio = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_linea_sin_costo.rpt")
               End If
               If opt_solo_costo = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_linea_sin_precio.rpt")
               End If
               If opt_ambos = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_linea.rpt")
               End If
               VAR_CADENA_FILTRO = "{VW_EXISTENCIAS_GENERALES_LINEA.INTE_EXI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_EXISTENCIAS_GENERALES_LINEA.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_EXISTENCIAS_GENERALES_LINEA.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
               If Me.opt_filtrar_ceros = True Then
                  If opt_solo_precio = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} >= 0"
                  End If
                  If opt_solo_costo = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} >= 0"
                  End If
                  If opt_ambos = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} >= 0"
                  End If
               End If
               If Me.opt_filtrar_negativos = True Then
                  If opt_solo_precio = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} > 0"
                  End If
                  If opt_solo_costo = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} > 0"
                  End If
                  If opt_ambos = True Then
                     VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} > 0"
                  End If
               End If
               VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and " + var_filtro
               reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de existencias por linea"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  If opt_solo_precio = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_linea_sin_costo_excel.rpt")
                  End If
                  If opt_solo_costo = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_linea_sin_precio_excel.rpt")
                  End If
                  If opt_ambos = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_existencias_generales_linea_excel.rpt")
                  End If
                  VAR_CADENA_FILTRO = "{VW_EXISTENCIAS_GENERALES_LINEA.INTE_EXI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_EXISTENCIAS_GENERALES_LINEA.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_EXISTENCIAS_GENERALES_LINEA.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                  If Me.opt_filtrar_ceros = True Then
                     If opt_solo_precio = True Then
                        VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} >= 0"
                     End If
                     If opt_solo_costo = True Then
                        VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} >= 0"
                     End If
                     If opt_ambos = True Then
                        VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} >= 0"
                     End If
                  End If
                  If Me.opt_filtrar_negativos = True Then
                     If opt_solo_precio = True Then
                        VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} > 0"
                     End If
                     If opt_solo_costo = True Then
                        VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} > 0"
                     End If
                     If opt_ambos = True Then
                        VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " AND {VW_EXISTENCIAS_GENERALES_LINEA.CANTIDAD} > 0"
                     End If
                  End If
                  VAR_CADENA_FILTRO = VAR_CADENA_FILTRO + " and " + var_filtro
                  reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_existencias" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            End If
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ALMACENES where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ARTICULOS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ENTRADAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_SALIDAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No se a seleccionado ningun almacén", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   If var_todos_articulos = 1 Then
   Else
        var_todos_articulos = 0
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(2) = "*" Then
         lv_articulos.selectedItem.SubItems(2) = ""
         lv_articulos.ListItems.item(i).Bold = False
         lv_articulos.ListItems.item(i).ForeColor = &H80000012
         lv_articulos.ListItems.item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.item(i).Bold = True
         lv_articulos.ListItems.item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   var_todos_articulos = 0
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(2) = "*" Then
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.item(i).Bold = False
      lv_articulos.ListItems.item(i).ForeColor = &H80000012
      lv_articulos.ListItems.item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(2) = "*"
      lv_articulos.ListItems.item(i).Bold = True
      lv_articulos.ListItems.item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_articulos.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   var_todos_articulos = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.item(i).Bold = False
      lv_articulos.ListItems.item(i).ForeColor = &H80000012
      lv_articulos.ListItems.item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   If var_todos_articulos = 1 Then
   Else
         var_todos_articulos = 0
   End If
   n = lv_articulos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_articulos.ListItems.item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.item(i).Bold = True
         lv_articulos.ListItems.item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_articulos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   var_todos_articulos = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_articulos.ListItems.Count
   For i = 1 To n
       lv_articulos.ListItems.item(i).SubItems(2) = "*"
       lv_articulos.ListItems.item(i).Bold = True
       lv_articulos.ListItems.item(i).ForeColor = &HFF0000
       lv_articulos.ListItems.item(i).ListSubItems(1).Bold = True
       lv_articulos.ListItems.item(i).ListSubItems(2).Bold = True
       lv_articulos.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_articulos.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
   Next
   lv_articulos.Refresh
End Sub

Private Sub Command1_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
       lv_almacenes.ListItems.item(i).SubItems(2) = "*"
       lv_almacenes.ListItems.item(i).Bold = True
       lv_almacenes.ListItems.item(i).ForeColor = &H8000&
       lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = True
       lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = True
       lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H8000&
       lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H8000&
   Next
   lv_almacenes.Refresh
End Sub

Private Sub Command10_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_lineas.ListItems.Count
   For i = 1 To n
       lv_lineas.ListItems.item(i).SubItems(2) = "*"
       lv_lineas.ListItems.item(i).Bold = True
       lv_lineas.ListItems.item(i).ForeColor = &HFF0000
       lv_lineas.ListItems.item(i).ListSubItems(1).Bold = True
       lv_lineas.ListItems.item(i).ListSubItems(2).Bold = True
       lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
   Next
   lv_lineas.Refresh
End Sub

Private Sub Command11_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim si As Integer
   Dim var_existen_almacenes As Integer
   Dim var_cadena_linea As String
   If IsDate(txt_fecha) Then
      si = MsgBox("Se calcularan las existencias al dia " + Format(CDate(txt_fecha), "Long Date"), vbYesNo, "ATENCION")
      If si = 6 Then
         var_n = lv_almacenes.ListItems.Count
         var_existen_almacenes = 0
         For var_i = 1 To var_n
             lv_almacenes.ListItems.item(var_i).Selected = True
             If lv_almacenes.selectedItem.SubItems(2) = "*" Then
                var_existen_almacenes = var_existen_almacenes + 1
             End If
         Next var_i
         If var_existen_almacenes > 0 Then
            cnn.BeginTrans
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select max(inte_exi_consecutivo) from TB_TEMP_EXISTENCIAS_ALMACENES", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "delete from TB_TEMP_EXISTENCIAS_FISICO_CONTRA_MOVIMIENTOS", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "Insert into TB_TEMP_EXISTENCIAS_ALMACENES (INTE_EXI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            cnn.CommandTimeout = 360
            If Me.opt_seleccion_articulos = True Then
               If Me.opt_todos = True Then
                  var_todos_articulos = 1
               Else
                  var_todos_articulos = 0
               End If
               If var_todos_articulos = 1 Then
                  rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
               Else
                  var_n = lv_articulos.ListItems.Count
                  For var_i = 1 To var_n
                      lv_articulos.ListItems.item(var_i).Selected = True
                      rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) values (" + CStr(var_consecutivo) + ", '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                  Next var_i
               End If
            End If
            
            If Me.opt_seleccion_linea = True Then
               If var_todos_lineas = 1 Then
                  If var_empresa = "18" Then
                     rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from VW_EXISTEN_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
                  End If
               Else
                  var_n = lv_lineas.ListItems.Count
                  For var_i = 1 To var_n
                      lv_lineas.ListItems.item(var_i).Selected = True
                       If Trim(lv_lineas.selectedItem.SubItems(2)) = "*" Then
                         rsaux.Open "SELECT VCHA_ART_ARTICULO_ID from tb_articulos where vcha_lin_linea_id = '" + lv_lineas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         While Not rsaux.EOF
                               rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) values (" + CStr(var_consecutivo) + ", '" + rsaux!VCHA_ART_ARTICULO_ID + "')", cnn, adOpenDynamic, adLockOptimistic
                               rsaux.MoveNext
                         Wend
                         rsaux.Close
                      End If
                  Next var_i
               End If
            End If
            
            var_fecha_fin_1 = CDate(txt_fecha) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            
             var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_año = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             var_fecha_fin_1 = CDate(txt_fecha)
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_año = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             'var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            
            var_n = lv_almacenes.ListItems.Count
            For var_i = 1 To var_n
                lv_almacenes.ListItems.item(var_i).Selected = True
                If lv_almacenes.selectedItem.SubItems(2) = "*" Then
                   rs.Open "Insert into TB_TEMP_EXISTENCIAS_ALMACENES (INTE_EXI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina, VCHA_ALM_ALMACEN_ID) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '" + lv_almacenes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            If rsaux5.State Then
               rsaux5.Close
            End If
            
            rsaux5.Open "select * from TB_TEMP_EXISTENCIAS_ALMACENES where inte_exi_consecutivo = " + CStr(var_consecutivo) + " and vcha_alm_almacen_ID is not null", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommandTimeout = 3000
            var_filtro = ""
            While Not rsaux5.EOF
                  var_almacen = rsaux5!VCHA_ALM_ALMACEN_ID
                  If opt_general = True Then
                     If var_filtro = "" Then
                        var_filtro = var_filtro + "({VW_EXISTENCIAS_GENERALES.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     Else
                        var_filtro = var_filtro + " or {VW_EXISTENCIAS_GENERALES.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     End If
                  Else
                     If var_filtro = "" Then
                        var_filtro = var_filtro + "({VW_EXISTENCIAS_GENERALES_LINEA.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     Else
                        var_filtro = var_filtro + " or {VW_EXISTENCIAS_GENERALES_LINEA.vcha_alm_almacen_id} = '" + var_almacen + "'"
                     End If
                  End If
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_1 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_2 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_3 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_4 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_5 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            var_filtro = var_filtro + ")"
            ' se quitaron el dia 26-09-2008
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_1 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_2 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_3 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_4 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_5 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
            
            'rs.Open "exec SP_EXISTENCIAS_RAPIDAS_6 " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
           
            rs.Open "exec SP_EXISTENCIAS_RAPIDAS_CUADRE " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin + ",'" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                         
                         
            If opt_general = True Then
               Set reporte = appl.OpenReport(App.Path + "\rep_existencias_fisico_contra_movimientos.rpt")
               frmvistasprevias.cr.ReportSource = reporte
               reporte.RecordSelectionFormula = "{VW_REPORTE_EXISTENCIAS_FISICO_CONTRA_MOVIMIENTOS.inte_tem_consecutivo} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de existencias"
               frmvistasprevias.Show 1
               Set reporte = Nothing
            
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_existencias_fisico_contra_movimient.rpt ")
                  frmvistasprevias.cr.ReportSource = reporte
                  reporte.RecordSelectionFormula = "{VW_REPORTE_EXISTENCIAS_FISICO_CONTRA_MOVIMIENTOS.inte_tem_consecutivo} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_existencias" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            End If
            If opt_linea = True Then
               reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de existencias por linea"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_existencias" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            End If
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ALMACENES where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ARTICULOS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ENTRADAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_SALIDAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_FISICO_CONTRA_MOVIMIENTOS where INTE_tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No se a seleccionado ningun almacén", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command12_Click()
   If IsDate(Me.txt_fecha) Then
      mes.Value = CDate(txt_fecha)
   Else
      mes.Value = Date
   End If
   mes.Visible = True
End Sub

Private Sub Command2_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
       lv_almacenes.ListItems.item(i).SubItems(2) = " "
       lv_almacenes.ListItems.item(i).Bold = False
       lv_almacenes.ListItems.item(i).ForeColor = &H80000012
       lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = False
       lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = False
       lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
       lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
    Next
    lv_almacenes.Refresh
End Sub

Private Sub Command3_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
       If lv_almacenes.ListItems.item(i).SubItems(2) = "*" Then
          lv_almacenes.ListItems.item(i).SubItems(2) = " "
          lv_almacenes.ListItems.item(i).Bold = False
          lv_almacenes.ListItems.item(i).ForeColor = &H80000012
          lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = False
          lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = False
          lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
          lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_almacenes.ListItems.item(i).SubItems(2) = "*"
          lv_almacenes.ListItems.item(i).Bold = True
          lv_almacenes.ListItems.item(i).ForeColor = &H8000&
          lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = True
          lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = True
          lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H8000&
          lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H8000&
      End If
   Next
   lv_almacenes.Refresh
End Sub

Private Sub Command4_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   i = lv_almacenes.selectedItem.Index
   If lv_almacenes.selectedItem.SubItems(2) = "*" Then
      lv_almacenes.ListItems.item(i).SubItems(2) = " "
      lv_almacenes.ListItems.item(i).Bold = False
      lv_almacenes.ListItems.item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_almacenes.ListItems.item(i).SubItems(2) = "*"
      lv_almacenes.ListItems.item(i).Bold = True
      lv_almacenes.ListItems.item(i).ForeColor = &H8000&
      lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = True
      lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = True
      lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H8000&
      lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H8000&
   End If
   lv_almacenes.Refresh
End Sub

Private Sub Command5_Click()
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
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
       If lv_almacenes.ListItems.item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_almacenes.ListItems.item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_almacenes.ListItems.item(i).SubItems(2) = "*"
       lv_almacenes.ListItems.item(i).Bold = True
       lv_almacenes.ListItems.item(i).ForeColor = &H8000&
       lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = True
       lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = True
       lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H8000&
       lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H8000&
       lv_almacenes.Refresh
   Next
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_lineas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_lineas.ListItems.item(i).Selected = True
      If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.item(i).Bold = True
         lv_lineas.ListItems.item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_lineas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_lineas.selectedItem.Index
   If lv_lineas.selectedItem.SubItems(2) = "*" Then
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.item(i).Bold = False
      lv_lineas.ListItems.item(i).ForeColor = &H80000012
      lv_lineas.ListItems.item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lineas.Refresh
   Else
      lv_lineas.selectedItem.SubItems(2) = "*"
      lv_lineas.ListItems.item(i).Bold = True
      lv_lineas.ListItems.item(i).ForeColor = &HFF0000
      lv_lineas.ListItems.item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_lineas.Refresh
   End If
End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.item(i).Selected = True
      If lv_lineas.selectedItem.SubItems(2) = "*" Then
         lv_lineas.selectedItem.SubItems(2) = ""
         lv_lineas.ListItems.item(i).Bold = False
         lv_lineas.ListItems.item(i).ForeColor = &H80000012
         lv_lineas.ListItems.item(i).ListSubItems(1).Bold = False
         lv_lineas.ListItems.item(i).ListSubItems(2).Bold = False
         lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.item(i).Bold = True
         lv_lineas.ListItems.item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.item(i).Selected = True
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.item(i).Bold = False
      lv_lineas.ListItems.item(i).ForeColor = &H80000012
      lv_lineas.ListItems.item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_lineas.Refresh
End Sub

Private Sub Form_Load()
   If var_clave_usuario_global = "8" Then
      Me.Command11.Visible = True
   Else
      Me.Command11.Visible = False
   End If
   Me.opt_seleccion = True
   Me.opt_filtrar_ninguna = True
   var_cadena_seguridad = ""
   var_todos_articulos = 0
   Top = 0
   Left = 0
   opt_general = True
   opt_solo_precio = True
   mes.Visible = False
   txt_fecha = Date
   Dim list_item As ListItem
   If var_empresa = "16" Or var_empresa = "31" Or var_empresa = "18" Or var_empresa = "06" Or var_empresa = "15" Then
      rs.Open "SELECT DISTINCT dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE FROM         dbo.TB_ALMACENES INNER JOIN    dbo.TB_EXISTENCIAS ON dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "SELECT DISTINCT dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE FROM         dbo.TB_ALMACENES INNER JOIN    dbo.TB_EXISTENCIAS ON dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID = dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID", cnn, adOpenDynamic, adLockOptimistic
   End If
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_almacenes.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext
   Wend
   rs.Close
   'rs.Open "select * from TB_ARTICULOS order by vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
   'While Not rs.EOF
   '    Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
   '    list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
   '    list_item.SubItems(2) = " "
   '    rs.MoveNext
   'Wend
   'rs.Close
   rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
       Set list_item = lv_lineas.ListItems.Add(, , rs!vcha_lin_linea_id)
       list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE))
       list_item.SubItems(2) = " "
       rs.MoveNext
   Wend
   rs.Close
   Me.opt_seleccion_articulos = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_almacenes.ListItems.Count
      i = lv_almacenes.selectedItem.Index
      If lv_almacenes.ListItems.item(i).SubItems(2) = "*" Then
      lv_almacenes.ListItems.item(i).SubItems(2) = " "
             lv_almacenes.ListItems.item(i).Bold = False
             lv_almacenes.ListItems.item(i).ForeColor = &H80000012
             lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = False
             lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = False
             lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
             lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
          Else
             lv_almacenes.ListItems.item(i).SubItems(2) = "*"
             lv_almacenes.ListItems.item(i).Bold = True
             lv_almacenes.ListItems.item(i).ForeColor = &H8000&
             lv_almacenes.ListItems.item(i).ListSubItems(1).Bold = True
             lv_almacenes.ListItems.item(i).ListSubItems(2).Bold = True
             lv_almacenes.ListItems.item(i).ListSubItems(1).ForeColor = &H8000&
             lv_almacenes.ListItems.item(i).ListSubItems(2).ForeColor = &H8000&
         End If
      lv_almacenes.Refresh
   End If
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   txt_fecha = mes.Value
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fecha = mes.Value
      mes.Visible = False
   End If
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub opt_seleccion_articulos_Click()
   Me.frm_linea.Visible = False
   Me.frm_Articulo.Visible = True
End Sub

Private Sub opt_seleccion_linea_Click()
   Me.frm_Articulo.Visible = False
   Me.frm_linea.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_lineas, Text1, False)
      Text1 = ""
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim itmfound As ListItem
      var_posible = False
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
         rs.Close
      Else
         rs.Close
         rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_posible = True
               txt_buscar = rs!VCHA_ART_ARTICULO_ID
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
      rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_codigo = txt_buscar
         valor = var_codigo
        If Trim(var_codigo) <> "" Then
           Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
         Else
           Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
         End If
    
         If itmfound Is Nothing Then
            Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
            If itmfound Is Nothing Then
               Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
               If itmfound Is Nothing Then
                  Set list_item = lv_articulos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                  list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español))
                  txt_buscar = ""
                  Exit Sub
               Else
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  lv_articulos.SetFocus
                End If
            Else
               itmfound.EnsureVisible
               itmfound.Selected = True
               lv_articulos.SetFocus
            End If
         Else
            itmfound.EnsureVisible
            itmfound.Selected = True
            lv_articulos.SetFocus
         End If
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
      txt_buscar = ""
   End If
End Sub

