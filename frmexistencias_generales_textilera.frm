VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmexistencias_generales_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Existencias Generales Textilera"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmexistencias_generales_textilera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   330
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1395
      TabIndex        =   14
      Top             =   4815
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   69337089
      CurrentDate     =   37761
   End
   Begin VB.Frame frm_Articulo 
      Caption         =   " Art�culos "
      Height          =   6885
      Left            =   5730
      TabIndex        =   29
      Top             =   390
      Width           =   5865
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2205
         TabIndex        =   38
         Top             =   750
         Width           =   1740
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmexistencias_generales_textilera.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   105
         Picture         =   "frmexistencias_generales_textilera.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmexistencias_generales_textilera.frx":041A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Invertir Selecci�n Alt + V"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmexistencias_generales_textilera.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Marcar (Enter)"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmexistencias_generales_textilera.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   525
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame Frame8 
         Height          =   120
         Left            =   15
         TabIndex        =   32
         Top             =   555
         Width           =   5820
      End
      Begin VB.OptionButton opt_todos 
         Caption         =   "Todos los art�culos"
         Height          =   225
         Left            =   3360
         TabIndex        =   31
         Top             =   255
         Width           =   1845
      End
      Begin VB.OptionButton opt_seleccion 
         Caption         =   "Selecci�n de art�culos"
         Height          =   345
         Left            =   435
         TabIndex        =   30
         Top             =   240
         Width           =   1980
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   5700
         Left            =   60
         TabIndex        =   39
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
            Text            =   "C�digo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci�n"
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
         Caption         =   "Busqueda de art�culo:"
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   40
         Top             =   810
         Width           =   1575
      End
   End
   Begin VB.Frame frm_canales 
      Height          =   6105
      Left            =   90
      TabIndex        =   21
      Top             =   390
      Width           =   5595
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmexistencias_generales_textilera.frx":094C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmexistencias_generales_textilera.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Marcar (Enter)"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmexistencias_generales_textilera.frx":0DAC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Invertir Selecci�n Alt + V"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   30
         Picture         =   "frmexistencias_generales_textilera.frx":0E7E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmexistencias_generales_textilera.frx":0F80
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   360
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   5370
         Left            =   45
         TabIndex        =   27
         Top             =   690
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   9472
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
         TabIndex        =   28
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmexistencias_generales_textilera.frx":1196
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   19
      Top             =   315
      Width           =   11640
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11235
      Picture         =   "frmexistencias_generales_textilera.frx":1298
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Fecha "
      Height          =   720
      Left            =   90
      TabIndex        =   15
      Top             =   6555
      Width           =   5595
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   255
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3105
         Picture         =   "frmexistencias_generales_textilera.frx":18D2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar Reporte Agrupado Por "
      Height          =   1110
      Left            =   75
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   2715
      Begin VB.OptionButton opt_general 
         Caption         =   "Art�culo"
         Height          =   270
         Left            =   210
         TabIndex        =   13
         Top             =   240
         Width           =   1200
      End
      Begin VB.OptionButton opt_linea 
         Caption         =   "Linea"
         Height          =   270
         Left            =   210
         TabIndex        =   12
         Top             =   555
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Filtrar Reporte Por "
      Height          =   1110
      Left            =   75
      TabIndex        =   7
      Top             =   390
      Visible         =   0   'False
      Width           =   5625
      Begin VB.OptionButton opt_solo_precio 
         Caption         =   "Solo Precio"
         Height          =   240
         Left            =   165
         TabIndex        =   10
         Top             =   195
         Width           =   2325
      End
      Begin VB.OptionButton opt_solo_costo 
         Caption         =   "Solo Costo"
         Height          =   240
         Left            =   165
         TabIndex        =   9
         Top             =   495
         Width           =   2325
      End
      Begin VB.OptionButton opt_ambos 
         Caption         =   "Ambos"
         Height          =   240
         Left            =   165
         TabIndex        =   8
         Top             =   795
         Width           =   2325
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Filtrar Reporte Por "
      Height          =   1110
      Left            =   2955
      TabIndex        =   3
      Top             =   1500
      Visible         =   0   'False
      Width           =   2715
      Begin VB.OptionButton opt_filtrar_ceros 
         Caption         =   "Existencias en cero"
         Height          =   270
         Left            =   210
         TabIndex        =   6
         Top             =   510
         Width           =   2340
      End
      Begin VB.OptionButton opt_filtrar_negativos 
         Caption         =   "Negativos"
         Height          =   270
         Left            =   210
         TabIndex        =   5
         Top             =   780
         Width           =   1140
      End
      Begin VB.OptionButton opt_filtrar_ninguna 
         Caption         =   "Ninguna"
         Height          =   270
         Left            =   255
         TabIndex        =   4
         Top             =   225
         Width           =   2340
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Seleccionar "
      Height          =   1110
      Left            =   90
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   2715
      Begin VB.OptionButton opt_seleccion_articulos 
         Caption         =   "Art�culo"
         Height          =   270
         Left            =   210
         TabIndex        =   2
         Top             =   330
         Width           =   1200
      End
      Begin VB.OptionButton opt_seleccion_linea 
         Caption         =   "Linea"
         Height          =   270
         Left            =   210
         TabIndex        =   1
         Top             =   645
         Width           =   1140
      End
   End
   Begin VB.Frame frm_linea 
      Caption         =   " Lineas "
      Height          =   6885
      Left            =   5730
      TabIndex        =   41
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   15
         TabIndex        =   48
         Top             =   525
         Width           =   5820
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmexistencias_generales_textilera.frx":2B44
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmexistencias_generales_textilera.frx":2D5A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmexistencias_generales_textilera.frx":2FA4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Invertir Selecci�n Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   105
         Picture         =   "frmexistencias_generales_textilera.frx":3076
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmexistencias_generales_textilera.frx":3178
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1815
         TabIndex        =   42
         Top             =   780
         Width           =   1620
      End
      Begin MSComctlLib.ListView lv_lineas 
         Height          =   5625
         Left            =   60
         TabIndex        =   49
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
            Text            =   "C�digo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci�n"
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
         TabIndex        =   50
         Top             =   795
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmexistencias_generales_textilera"
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
   Dim a�o As String
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
             lv_almacenes.ListItems.Item(var_i).Selected = True
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
                      lv_articulos.ListItems.Item(var_i).Selected = True
                      rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + ", vcha_Art_articulo_id from tb_articulos where vcha_art_articulo_id like '" + Me.lv_articulos.selectedItem + "%'", cnn, adOpenDynamic, adLockOptimistic
                  Next var_i
               End If
            End If
            
            If Me.opt_seleccion_linea = True Then
               If var_todos_lineas = 1 Then
                  rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) select " + CStr(var_consecutivo) + " as numero, vcha_Art_articulo_id from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
               Else
                  var_n = lv_lineas.ListItems.Count
                  For var_i = 1 To var_n
                      lv_lineas.ListItems.Item(var_i).Selected = True
                       If Trim(lv_lineas.selectedItem.SubItems(2)) = "*" Then
                         rsaux.Open "SELECT VCHA_ART_ARTICULO_ID from tb_articulos where vcha_lin_linea_id = '" + lv_lineas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         While Not rsaux.EOF
                               rs.Open "insert into tb_temp_existencias_articulos (inte_Exi_consecutivo, vcha_Art_articulo_id) values (" + CStr(var_consecutivo) + ", '" + rsaux!vcha_Art_articulo_id + "')", cnn, adOpenDynamic, adLockOptimistic
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
            var_a�o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            
             var_fecha_inicio = "{d '" + var_a�o + "-" + var_mes + "-" + var_dia + "'}"
             
             
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_a�o = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_a�o + "-" + var_mes + "-" + var_dia + "'}"
             
             var_fecha_fin_1 = CDate(txt_fecha)
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_a�o = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             'var_fecha_fin = "{d '" + var_a�o + "-" + var_mes + "-" + var_dia + "'}"
             
            
            var_n = lv_almacenes.ListItems.Count
            For var_i = 1 To var_n
                lv_almacenes.ListItems.Item(var_i).Selected = True
                If lv_almacenes.selectedItem.SubItems(2) = "*" Then
                   rs.Open "Insert into TB_TEMP_EXISTENCIAS_ALMACENES (INTE_EXI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina, VCHA_ALM_ALMACEN_ID) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '" + lv_almacenes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            
            rs.Open "exec SP_EXISTENCIAS_RAPIDAS_TEXTILERA " + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
           
            Set reporte = appl.OpenReport(App.Path + "\rep_existencias_TEXTILERA.rpt")
            VAR_CADENA_FILTRO = "{VW_EXISTENCIAS_TEXTILERA.INTE_EXI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_EXISTENCIAS_TEXTILERA.VCHA_AUD_USUARIO}= '" + var_clave_usuario_global + "' and {VW_EXISTENCIAS_TEXTILERA.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
            reporte.RecordSelectionFormula = VAR_CADENA_FILTRO
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de existencias"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("�Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_existencias_TEXTILERA.rpt")
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
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ALMACENES where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ARTICULOS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_ENTRADAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS_SALIDAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_EXISTENCIAS where INTE_EXI_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No se a seleccionado ningun almac�n", vbOKOnly, "ATENCION"
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
      lv_articulos.ListItems.Item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(2) = "*" Then
         lv_articulos.selectedItem.SubItems(2) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   var_todos_articulos = 0
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(2) = "*" Then
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(2) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_articulos.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   var_todos_articulos = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
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
      lv_articulos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
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
       lv_articulos.ListItems.Item(i).SubItems(2) = "*"
       lv_articulos.ListItems.Item(i).Bold = True
       lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
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
       lv_almacenes.ListItems.Item(i).SubItems(2) = "*"
       lv_almacenes.ListItems.Item(i).Bold = True
       lv_almacenes.ListItems.Item(i).ForeColor = &H8000&
       lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
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
       lv_lineas.ListItems.Item(i).SubItems(2) = "*"
       lv_lineas.ListItems.Item(i).Bold = True
       lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
       lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next
   lv_lineas.Refresh
End Sub

Private Sub Command11_Click()
   Me.opt_seleccion = True
   Me.opt_filtrar_ninguna = True
   var_cadena_seguridad = ""
   var_todos_articulos = 0
   opt_general = True
   opt_solo_precio = True
   mes.Visible = False
   txt_fecha = Date
   Dim list_item As ListItem
   Me.lv_almacenes.ListItems.Clear
   Me.lv_articulos.ListItems.Clear
   rs.Open "select * from tb_almacenes order by vcha_Alm_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_almacenes.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext
   Wend
   rs.Close
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
       lv_almacenes.ListItems.Item(i).SubItems(2) = " "
       lv_almacenes.ListItems.Item(i).Bold = False
       lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
       lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
       lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
       lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
       lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
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
       If lv_almacenes.ListItems.Item(i).SubItems(2) = "*" Then
          lv_almacenes.ListItems.Item(i).SubItems(2) = " "
          lv_almacenes.ListItems.Item(i).Bold = False
          lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
          lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
       Else
          lv_almacenes.ListItems.Item(i).SubItems(2) = "*"
          lv_almacenes.ListItems.Item(i).Bold = True
          lv_almacenes.ListItems.Item(i).ForeColor = &H8000&
          lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
          lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
          lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
          lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
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
      lv_almacenes.ListItems.Item(i).SubItems(2) = " "
      lv_almacenes.ListItems.Item(i).Bold = False
      lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Else
      lv_almacenes.ListItems.Item(i).SubItems(2) = "*"
      lv_almacenes.ListItems.Item(i).Bold = True
      lv_almacenes.ListItems.Item(i).ForeColor = &H8000&
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
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
       If lv_almacenes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
          numero_seleccionado1 = i
          primera_vez = True
       End If
       If lv_almacenes.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
          numero_seleccionado2 = i
       End If
   Next
   For i = numero_seleccionado1 To numero_seleccionado2
       lv_almacenes.ListItems.Item(i).SubItems(2) = "*"
       lv_almacenes.ListItems.Item(i).Bold = True
       lv_almacenes.ListItems.Item(i).ForeColor = &H8000&
       lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
       lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
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
      lv_lineas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_lineas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
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
      lv_lineas.ListItems.Item(i).Bold = False
      lv_lineas.ListItems.Item(i).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lineas.Refresh
   Else
      lv_lineas.selectedItem.SubItems(2) = "*"
      lv_lineas.ListItems.Item(i).Bold = True
      lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
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
      lv_lineas.ListItems.Item(i).Selected = True
      If lv_lineas.selectedItem.SubItems(2) = "*" Then
         lv_lineas.selectedItem.SubItems(2) = ""
         lv_lineas.ListItems.Item(i).Bold = False
         lv_lineas.ListItems.Item(i).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_lineas.selectedItem.SubItems(2) = "*"
         lv_lineas.ListItems.Item(i).Bold = True
         lv_lineas.ListItems.Item(i).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_lineas.ListItems.Count
   For i = 1 To n
      lv_lineas.ListItems.Item(i).Selected = True
      lv_lineas.selectedItem.SubItems(2) = ""
      lv_lineas.ListItems.Item(i).Bold = False
      lv_lineas.ListItems.Item(i).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lineas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lineas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_lineas.Refresh
End Sub

Private Sub Form_Load()
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
   rs.Open "select * from tb_almacenes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_Alm_nombre ", cnn, adOpenDynamic, adLockOptimistic
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
   '    list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_art_nombre_espa�ol), "", rs!vcha_art_nombre_espa�ol))
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
      If lv_almacenes.ListItems.Item(i).SubItems(2) = "*" Then
      lv_almacenes.ListItems.Item(i).SubItems(2) = " "
             lv_almacenes.ListItems.Item(i).Bold = False
             lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
             lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
             lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
             lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
             lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
          Else
             lv_almacenes.ListItems.Item(i).SubItems(2) = "*"
             lv_almacenes.ListItems.Item(i).Bold = True
             lv_almacenes.ListItems.Item(i).ForeColor = &H8000&
             lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
             lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
             lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
             lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
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
      If Len(Trim(Me.txt_buscar)) = 1 Then
         rs.Open "select * from tb_tipos_productos where vcha_tpr_tipo_producto_id = '" + Me.txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_DEscripcion = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
            var_codigo = txt_buscar
            valor = var_codigo
            If var_codigo Then
               Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
            Else
               Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
            End If
    
            If itmfound Is Nothing Then
               Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
               If itmfound Is Nothing Then
                  Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                  If itmfound Is Nothing Then
                     Set list_item = lv_articulos.ListItems.Add(, , txt_buscar)
                     list_item.SubItems(1) = var_DEscripcion
                     list_item.SubItems(2) = "*"
                     txt_buscar = ""
                     rs.Close
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
            MsgBox "Tipo de producto no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         If Len(Trim(Me.txt_buscar)) = 3 Then
            var_tipo = Left(Me.txt_buscar, 1)
            var_division = Mid(Me.txt_buscar, 2, 2)
            rs.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + var_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + var_division + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_DEscripcion = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
               var_codigo = Me.txt_buscar
               valor = var_codigo
               If var_codigo Then
                  Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
               Else
                  Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
               End If
  
               If itmfound Is Nothing Then
                  Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                  If itmfound Is Nothing Then
                     Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                     If itmfound Is Nothing Then
                        Set list_item = lv_articulos.ListItems.Add(, , txt_buscar)
                        list_item.SubItems(1) = var_DEscripcion
                        list_item.SubItems(2) = "*"
                        txt_buscar = ""
                        rs.Close
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
               MsgBox "La division no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            If Len(Trim(Me.txt_buscar)) = 5 Then
               var_tipo = Left(Me.txt_buscar, 1)
               var_division = Mid(Me.txt_buscar, 2, 2)
               var_subdivision = Mid(Me.txt_buscar, 4, 2)
               rs.Open "SELECT * FROM TB_subDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + var_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + var_division + "' and vcha_sub_subdivision_id = '" + var_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_DEscripcion = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
                  var_codigo = Me.txt_buscar
                  valor = var_codigo
                  If var_codigo Then
                     Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                  Else
                     Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                  End If
  
                  If itmfound Is Nothing Then
                     Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                     If itmfound Is Nothing Then
                        Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                        If itmfound Is Nothing Then
                           Set list_item = lv_articulos.ListItems.Add(, , txt_buscar)
                           list_item.SubItems(1) = var_DEscripcion
                           list_item.SubItems(2) = "*"
                           txt_buscar = ""
                           rs.Close
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
                  MsgBox "La subdivision no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
            
            
            
            
            
            
            
            Else
               If Len(Trim(Me.txt_buscar)) = 10 Then
               
               
                  
                  var_tipo = Left(Me.txt_buscar, 1)
                  var_division = Mid(Me.txt_buscar, 2, 2)
                  var_subdivision = Mid(Me.txt_buscar, 4, 2)
                  VAR_ESTAMPADO = Mid(Me.txt_buscar, 6, 5)
                  rs.Open "SELECT * FROM TB_subDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + var_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + var_division + "' and vcha_sub_subdivision_id = '" + var_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_EST_ESTAMPADO_ID = '" + VAR_ESTAMPADO + "'", cnn, adOpenDynamic, adLockOptimistic
                     VAR_ESTAMPADO = ""
                     If Not rsaux.EOF Then
                        VAR_ESTAMPADO = IIf(IsNull(rsaux!vcha_est_nombre), "", rsaux!vcha_est_nombre)
                     End If
                     rsaux.Close
                     var_DEscripcion = Trim(IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)) + " " + VAR_ESTAMPADO
                     var_codigo = Me.txt_buscar
                     valor = var_codigo
                     If var_codigo Then
                        Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                     Else
                        Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                     End If
     
                     If itmfound Is Nothing Then
                        Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                        If itmfound Is Nothing Then
                           Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                           If itmfound Is Nothing Then
                              Set list_item = lv_articulos.ListItems.Add(, , txt_buscar)
                              list_item.SubItems(1) = var_DEscripcion
                              list_item.SubItems(2) = "*"
                              txt_buscar = ""
                              rs.Close
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
                     MsgBox "La subdivision no existe", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
               
               
               
               
               
               Else
                  MsgBox "C�digo incorrecto", vbOKCancel, "ATENCION"
               End If
            End If
         End If
      End If
   End If
End Sub


