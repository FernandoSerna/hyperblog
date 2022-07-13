VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreporte_entradas_produccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de entradas de producción"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   " Artículos "
      Height          =   6795
      Left            =   5760
      TabIndex        =   28
      Top             =   420
      Width           =   5865
      Begin VB.TextBox txt_buscar 
         Height          =   285
         Left            =   1800
         TabIndex        =   36
         Top             =   765
         Width           =   1350
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmreporte_entradas_produccion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   105
         Picture         =   "frmreporte_entradas_produccion.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmreporte_entradas_produccion.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmreporte_entradas_produccion.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmreporte_entradas_produccion.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   5565
         Left            =   60
         TabIndex        =   34
         Top             =   1170
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   9816
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
      Begin VB.Frame Frame8 
         Height          =   120
         Left            =   15
         TabIndex        =   35
         Top             =   525
         Width           =   5820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de artículo:"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   37
         Top             =   795
         Width           =   1575
      End
   End
   Begin MSComCtl2.MonthView mon_mes2 
      Height          =   2370
      Left            =   2745
      TabIndex        =   0
      Top             =   4455
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   59834369
      CurrentDate     =   37761
   End
   Begin MSComCtl2.MonthView mon_mes1 
      Height          =   2370
      Left            =   135
      TabIndex        =   14
      Top             =   4440
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   59834369
      CurrentDate     =   37761
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11325
      Picture         =   "frmreporte_entradas_produccion.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_entradas_produccion.frx":0E84
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   0
      TabIndex        =   27
      Top             =   360
      Width           =   11715
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Plantas "
      Height          =   3585
      Left            =   60
      TabIndex        =   24
      Top             =   2895
      Width           =   5685
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   15
         TabIndex        =   25
         Top             =   525
         Width           =   5640
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmreporte_entradas_produccion.frx":0F86
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   105
         Picture         =   "frmreporte_entradas_produccion.frx":119C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmreporte_entradas_produccion.frx":129E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmreporte_entradas_produccion.frx":1370
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmreporte_entradas_produccion.frx":15BA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_plantas 
         Height          =   2835
         Left            =   30
         TabIndex        =   26
         Top             =   675
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   5001
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
            Text            =   "Unidad"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   2083
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5734
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Movimientos "
      Height          =   2445
      Left            =   60
      TabIndex        =   22
      Top             =   420
      Width           =   5685
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmreporte_entradas_produccion.frx":17D0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmreporte_entradas_produccion.frx":19E6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmreporte_entradas_produccion.frx":1C30
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   105
         Picture         =   "frmreporte_entradas_produccion.frx":1D02
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmreporte_entradas_produccion.frx":1E04
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   15
         TabIndex        =   23
         Top             =   525
         Width           =   5640
      End
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   1620
         Left            =   15
         TabIndex        =   8
         Top             =   720
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2858
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   90
      TabIndex        =   15
      Top             =   6495
      Width           =   5670
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3585
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   293
         Width           =   1305
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   293
         Width           =   1320
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2355
         Picture         =   "frmreporte_entradas_produccion.frx":201A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fecha Inicial"
         Top             =   300
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4920
         Picture         =   "frmreporte_entradas_produccion.frx":328C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fecha Final"
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3270
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   20
         Top             =   353
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmreporte_entradas_produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_todos_articulos As Double

Private Sub cmd_imprimir_Click()
   Dim var_n As Double, var_i As Double, var_j As Double
   Dim var_contador_movimientos As Double, var_contador_plantas As Double
   Dim var_cadena_almacenes As String
   Dim var_cadena_movimientos As String
   Dim var_consecutivo As Double
   Dim var_cadena_articulos
   var_n = lv_movimientos.ListItems.Count
   var_contador_movimientos = 0
   If var_n > 0 Then
      For var_i = 1 To var_n
          lv_movimientos.ListItems.Item(var_i).Selected = True
          If Trim(lv_movimientos.selectedItem.SubItems(2)) = "*" Then
             var_contador_movimientos = var_contador_movimientos + 1
          End If
      Next var_i
      If var_contador_movimientos > 0 Then
         var_n = lv_plantas.ListItems.Count
         If var_n > 0 Then
            var_contador_plantas = 0
            For var_i = 1 To var_n
                lv_plantas.ListItems.Item(var_i).Selected = True
                If Trim(lv_plantas.selectedItem.SubItems(2)) = "*" Then
                   var_contador_plantas = var_contador_plantas + 1
                End If
            Next var_i
            If var_contador_plantas > 0 Then
               If IsDate(txt_inicio) Then
                  If IsDate(txt_fin) Then
                     If CDate(txt_inicio) <= CDate(txt_fin) Then
                        var_fecha_fin_1 = CDate(txt_fin) + 1
                        var_dia = CStr(Day(CDate(txt_inicio)))
                        var_mes = CStr(Month(CDate(txt_inicio)))
                        var_año = CStr(Year(CDate(txt_inicio)))
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
             
                        
                        var_n = lv_movimientos.ListItems.Count
                        var_cadena_movimientos = ""
                        var_cadena_almacenes = ""
                        var_j = 0
                        For var_i = 1 To var_n
                            lv_movimientos.ListItems.Item(var_i).Selected = True
                            If Trim(lv_movimientos.selectedItem.SubItems(2)) = "*" Then
                               If var_j = 0 Then
                                  var_cadena_movimientos = " (vcha_mov_movimiento_id = '" + Trim(lv_movimientos.selectedItem) + "'"
                                  var_j = 1
                               Else
                                  var_cadena_movimientos = var_cadena_movimientos + " or vcha_mov_movimiento_id = '" + Trim(lv_movimientos.selectedItem) + "'"
                               End If
                            End If
                        Next var_i
                        var_cadena_movimientos = var_cadena_movimientos + ")"
                        var_n = lv_plantas.ListItems.Count
                        var_j = 0
                        For var_i = 1 To var_n
                            lv_plantas.ListItems.Item(var_i).Selected = True
                            If Trim(lv_plantas.selectedItem.SubItems(2)) = "*" Then
                               If var_j = 0 Then
                                  var_cadena_almacenes = "(vcha_pro_proveedor_id = '" + Trim(lv_plantas.selectedItem) + "'"
                                  var_j = 1
                               Else
                                  var_cadena_almacenes = var_cadena_almacenes + " or vcha_pro_proveedor_id = '" + Trim(lv_plantas.selectedItem) + "'"
                               End If
                            End If
                        Next var_i
                        
                        var_cadena_almacenes = var_cadena_almacenes + ")"
                        
                        var_cadena = "select * from tb_encabezado_movimientos where " + var_cadena_movimientos + " and " + var_cadena_almacenes + " And (dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.01) AND CHAR_EMO_ESTATUS <> 'C'"
                        var_fecha_fin_1 = CDate(txt_fin)
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
                        cnn.BeginTrans
                        rs.Open "select max(isnull(INTE_TEM_CONSECUTIVO,0)) as numero from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
                        Else
                           var_consecutivo = 0
                        End If
                        var_consecutivo = var_consecutivo + 1
                        rs.Close
                        
                        rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTM_TEM_FECHA_FINAL, VCHA_AUR_MAQUINA, VCHA_AUD_USUARIO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + fun_NombrePc + "','" + var_clave_usuario_global + "', '','', '', '',0)", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
                        
                        If var_todos_articulos = 0 Then
                           var_n = lv_articulos.ListItems.Count
                           For var_i = 1 To var_n
                               lv_articulos.ListItems.Item(var_i).Selected = True
                               If Trim(lv_articulos.selectedItem.SubItems(2)) = "*" Then
                                  rs.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ",'" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                               End If
                           Next var_i
                        Else
                           rs.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) select " + CStr(var_consecutivo) + " AS CONSECUTIVO, vcha_Art_Articulo_id from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        'MsgBox var_cadena
                        rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           While Not rs.EOF
                                 rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTM_TEM_FECHA_FINAL, VCHA_AUR_MAQUINA, VCHA_AUD_USUARIO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + fun_NombrePc + "','" + var_clave_usuario_global + "', '" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_EMO_NUMERO) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 rs.MoveNext
                           Wend
                           Set reporte = appl.OpenReport(App.Path + "\rep_entradas_produccion_detalle.rpt")
                           reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_PRODUCCION_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.EnableSearchExpertButton = True
                           frmvistasprevias.cr.EnableSelectExpertButton = True
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Reporte de entradas de producción"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                           var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              Set reporte = appl.OpenReport(App.Path + "\rep_entradas_produccion_detalle_EXCEL.rpt")
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_PRODUCCION_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                              reporte.ExportOptions.FormatType = crEFTExcel80
                              reporte.ExportOptions.DestinationType = crEDTDiskFile
                              archivo = "c:\reportessid\entradas_produccion" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                              reporte.ExportOptions.DiskFileName = archivo
                              reporte.Export False
                              Set reporte = Nothing
                              MsgBox "Se a terminado de guardar el archivo " + archivo
                           End If
                           
                           Set reporte = appl.OpenReport(App.Path + "\rep_entradas_produccion_totales.rpt")
                           reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_PRODUCCION_totales.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.EnableSearchExpertButton = True
                           frmvistasprevias.cr.EnableSelectExpertButton = True
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Reporte de entradas de producción"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                           var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              Set reporte = appl.OpenReport(App.Path + "\rep_entradas_produccion_totales_excel.rpt")
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_PRODUCCION_totales.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                              reporte.ExportOptions.FormatType = crEFTExcel80
                              reporte.ExportOptions.DestinationType = crEDTDiskFile
                              archivo = "c:\reportessid\entradas_produccion_totales_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                              reporte.ExportOptions.DiskFileName = archivo
                              reporte.Export False
                              Set reporte = Nothing
                              MsgBox "Se a terminado de guardar el archivo " + archivo
                           End If
                           
                        Else
                           MsgBox "No existen respuesta para la petición solicitada", vbOKOnly, "ATENCION"
                        End If
                        rs.Close
                        rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION_ARTICULOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     Else
                        MsgBox "La fecha de inicio no debe de ser mayor a la fecha final", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a seleccionado ninguna planta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No existen plantas", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado ningun movimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen movimientos de entradas de producción", vbOKOnly, "ATENCION"
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
   lv_movimientos.Refresh
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
   Dim numero_lineas As Double
   Dim numero_seleccionado1 As Double
   Dim numero_seleccionado2 As Double
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Double
   Dim n As Double
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
   n = lv_movimientos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_movimientos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_movimientos.selectedItem.SubItems(2) = "*"
         lv_movimientos.ListItems.Item(i).Bold = True
         lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_movimientos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_movimientos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command10_Click()
   n = lv_plantas.ListItems.Count
   For i = 1 To n
      lv_plantas.ListItems.Item(i).Selected = True
      lv_plantas.selectedItem.SubItems(3) = "*"
      lv_plantas.ListItems.Item(i).Bold = True
      lv_plantas.ListItems.Item(i).ForeColor = &HFF0000
      lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
   Next i
   lv_plantas.Refresh
End Sub

Private Sub Command11_Click()
   If IsDate(txt_inicio) Then
      mon_mes1.Value = CDate(txt_inicio)
   Else
      mon_mes1.Value = Date
   End If
   mon_mes1.Visible = True
   mon_mes1.SetFocus
End Sub

Private Sub Command12_Click()
   If IsDate(txt_fin) Then
      mon_mes2.Value = CDate(txt_fin)
   Else
      mon_mes2.Value = Date
   End If
   mon_mes2.Visible = True
   mon_mes2.SetFocus
End Sub

Private Sub Command13_Click()
   Dim var_n As Double, var_i As Double, var_j As Double
   Dim var_contador_movimientos As Double, var_contador_plantas As Double
   Dim var_cadena_almacenes As String
   Dim var_cadena_movimientos As String
   Dim var_consecutivo As Double
   Dim var_cadena_articulos
   var_n = lv_movimientos.ListItems.Count
   var_contador_movimientos = 0
   If var_n > 0 Then
      For var_i = 1 To var_n
          lv_movimientos.ListItems.Item(var_i).Selected = True
          If Trim(lv_movimientos.selectedItem.SubItems(2)) = "*" Then
             var_contador_movimientos = var_contador_movimientos + 1
          End If
      Next var_i
      If var_contador_movimientos > 0 Then
         var_n = lv_plantas.ListItems.Count
         If var_n > 0 Then
            var_contador_plantas = 0
            For var_i = 1 To var_n
                lv_plantas.ListItems.Item(var_i).Selected = True
                If Trim(lv_plantas.selectedItem.SubItems(2)) = "*" Then
                   var_contador_plantas = var_contador_plantas + 1
                End If
            Next var_i
            If var_contador_plantas > 0 Then
               If IsDate(txt_inicio) Then
                  If IsDate(txt_fin) Then
                     If CDate(txt_inicio) <= CDate(txt_fin) Then
                        var_fecha_fin_1 = CDate(txt_fin) + 1
                        var_dia = CStr(Day(CDate(txt_inicio)))
                        var_mes = CStr(Month(CDate(txt_inicio)))
                        var_año = CStr(Year(CDate(txt_inicio)))
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
             
                        
                        var_n = lv_movimientos.ListItems.Count
                        var_cadena_movimientos = ""
                        var_cadena_almacenes = ""
                        var_j = 0
                        For var_i = 1 To var_n
                            lv_movimientos.ListItems.Item(var_i).Selected = True
                            If Trim(lv_movimientos.selectedItem.SubItems(2)) = "*" Then
                               If var_j = 0 Then
                                  var_cadena_movimientos = " (vcha_mov_movimiento_id = '" + Trim(lv_movimientos.selectedItem) + "'"
                                  var_j = 1
                               Else
                                  var_cadena_movimientos = var_cadena_movimientos + " or vcha_mov_movimiento_id = '" + Trim(lv_movimientos.selectedItem) + "'"
                               End If
                            End If
                            var_cadena_movimientos = var_cadena_movimientos + ")"
                        Next var_i
                        var_n = lv_plantas.ListItems.Count
                        var_j = 0
                        For var_i = 1 To var_n
                            lv_plantas.ListItems.Item(var_i).Selected = True
                            If Trim(lv_plantas.selectedItem.SubItems(2)) = "*" Then
                               If var_j = 0 Then
                                  var_cadena_almacenes = "(vcha_pro_proveedor_id = '" + Trim(lv_plantas.selectedItem) + "'"
                                  var_j = 1
                               Else
                                  var_cadena_almacenes = var_cadena_almacenes + " or vcha_pro_proveedor_id = '" + Trim(lv_plantas.selectedItem) + "'"
                               End If
                            End If
                        Next var_i
                        
                        var_cadena_almacenes = var_cadena_almacenes + ")"
                        
                        var_cadena = "select * from tb_encabezado_movimientos where " + var_cadena_movimientos + " and " + var_cadena_almacenes + " And (dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.01) AND CHAR_EMO_ESTATUS <> 'C'"
                        var_fecha_fin_1 = CDate(txt_fin)
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
                        cnn.BeginTrans
                        rs.Open "select max(isnull(INTE_TEM_CONSECUTIVO,0)) as numero from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
                        Else
                           var_consecutivo = 0
                        End If
                        var_consecutivo = var_consecutivo + 1
                        rs.Close
                        
                        rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTM_TEM_FECHA_FINAL, VCHA_AUR_MAQUINA, VCHA_AUD_USUARIO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + fun_NombrePc + "','" + var_clave_usuario_global + "', '','', '', '',0)"
                        cnn.CommitTrans
                        
                        If var_todos_articulos = 0 Then
                           var_n = lv_articulos.ListItems.Count
                           For var_i = 1 To var_n
                               lv_articulos.ListItems.Item(var_i).Selected = True
                               If Trim(lv_articulos.selectedItem.SubItems(2)) = "*" Then
                                  rs.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) VALUES (" + CStr(var_consecutivo) + ",'" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                               End If
                           Next var_i
                        Else
                           rs.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION_ARTICULOS (INTE_TEM_CONSECUTIVO, VCHA_ART_ARTICULO_ID) select " + CStr(var_consecutivo) + " AS CONSECUTIVO, vcha_Art_Articulo_id from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           While Not rs.EOF
                                 rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTM_TEM_FECHA_FINAL, VCHA_AUR_MAQUINA, VCHA_AUD_USUARIO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-0.01, '" + fun_NombrePc + "','" + var_clave_usuario_global + "', '" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_EMO_NUMERO) + ")"
                                 rs.MoveNext
                           Wend
                           Set reporte = appl.OpenReport(App.Path + "\rep_entradas_produccion_detalle.rpt")
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_PRODUCCION_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                           reporte.ExportOptions.FormatType = crEFTExcel80
                           reporte.ExportOptions.DestinationType = crEDTDiskFile
                           reporte.ExportOptions.DiskFileName = "c:\entradas_produccion.xls"
                           reporte.Export False
                           Set reporte = Nothing
                        Else
                           MsgBox "No existen respuesta para la petición solicitada", vbOKOnly, "ATENCION"
                        End If
                        rs.Close
                        rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION_ARTICULOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     Else
                        MsgBox "La fecha de inicio no debe de ser mayor a la fecha final", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a seleccionado ninguna planta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No existen plantas", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado ningun movimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen movimientos de entradas de producción", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Command2_Click()
   i = lv_movimientos.selectedItem.Index
   If lv_movimientos.selectedItem.SubItems(2) = "*" Then
      lv_movimientos.selectedItem.SubItems(2) = ""
      lv_movimientos.ListItems.Item(i).Bold = False
      lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_movimientos.Refresh
   Else
      lv_movimientos.selectedItem.SubItems(2) = "*"
      lv_movimientos.ListItems.Item(i).Bold = True
      lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_movimientos.Refresh
   End If
End Sub

Private Sub Command3_Click()
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      If lv_movimientos.selectedItem.SubItems(2) = "*" Then
         lv_movimientos.selectedItem.SubItems(2) = ""
         lv_movimientos.ListItems.Item(i).Bold = False
         lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_movimientos.selectedItem.SubItems(2) = "*"
         lv_movimientos.ListItems.Item(i).Bold = True
         lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command4_Click()
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      lv_movimientos.selectedItem.SubItems(2) = ""
      lv_movimientos.ListItems.Item(i).Bold = False
      lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_movimientos.Refresh
End Sub

Private Sub Command5_Click()
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      lv_movimientos.selectedItem.SubItems(2) = "*"
      lv_movimientos.ListItems.Item(i).Bold = True
      lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_movimientos.Refresh
End Sub

Private Sub Command6_Click()
   n = lv_plantas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_plantas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_plantas.selectedItem.SubItems(3) = "" And var_rellena = True Then
         lv_plantas.selectedItem.SubItems(3) = "*"
         lv_plantas.ListItems.Item(i).Bold = True
         lv_plantas.ListItems.Item(i).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_plantas.selectedItem.SubItems(3) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_plantas.selectedItem.SubItems(3) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command7_Click()
   i = lv_plantas.selectedItem.Index
   If lv_plantas.selectedItem.SubItems(3) = "*" Then
      lv_plantas.selectedItem.SubItems(3) = ""
      lv_plantas.ListItems.Item(i).Bold = False
      lv_plantas.ListItems.Item(i).ForeColor = &H80000012
      lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_plantas.Refresh
   Else
      lv_plantas.selectedItem.SubItems(3) = "*"
      lv_plantas.ListItems.Item(i).Bold = True
      lv_plantas.ListItems.Item(i).ForeColor = &HFF0000
      lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_plantas.Refresh
   End If

End Sub

Private Sub Command8_Click()
   n = lv_plantas.ListItems.Count
   For i = 1 To n
      lv_plantas.ListItems.Item(i).Selected = True
      If lv_plantas.selectedItem.SubItems(3) = "*" Then
         lv_plantas.selectedItem.SubItems(3) = ""
         lv_plantas.ListItems.Item(i).Bold = False
         lv_plantas.ListItems.Item(i).ForeColor = &H80000012
         lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      Else
         lv_plantas.selectedItem.SubItems(3) = "*"
         lv_plantas.ListItems.Item(i).Bold = True
         lv_plantas.ListItems.Item(i).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   n = lv_plantas.ListItems.Count
   For i = 1 To n
      lv_plantas.ListItems.Item(i).Selected = True
      lv_plantas.selectedItem.SubItems(3) = ""
      lv_plantas.ListItems.Item(i).Bold = False
      lv_plantas.ListItems.Item(i).ForeColor = &H80000012
      lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
   Next i
   lv_plantas.Refresh
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_todos_articulos = 0
   txt_inicio = Date
   txt_fin = Date
   Top = 0
   Me.Left = 0
   mon_mes1.Visible = False
   mon_mes2.Visible = False
   rs.Open "select * from xxvia_tb_tipo_movimientos where numb_tmo_movimiento_id in (21,51,2,3,0) order by vcha_tmo_descripcion", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = Me.lv_movimientos.ListItems.Add(, , rs!numb_tmo_movimiento_id)
         list_item.SubItems(1) = UCase(rs!vcha_tmo_descripcion)
         var_contador = var_contador + 1
         rs.MoveNext
   Wend
   rs.Close
   If numero_items_movimientos > 12 Then
      lv_movimientos.ColumnHeaders(2).Width = 4200.71
   Else
      lv_movimientos.ColumnHeaders(2).Width = 4499.71
   End If

   rs.Open "SELECT organization_id, secondary_inventory_name, description  FROM mtl_secondary_inventories WHERE secondary_inventory_name not like 'CDI_TD%' and secondary_inventory_name not like 'TX_TD%'  order by organization_id, secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_plantas = 0
   While Not rs.EOF
      Set list_item = lv_plantas.ListItems.Add(, , rs!organization_id)
      list_item.SubItems(1) = IIf(IsNull(rs!secondary_inventory_name), "", rs!secondary_inventory_name)
      list_item.SubItems(2) = UCase(IIf(IsNull(rs!Description), "", rs!Description))
      list_item.SubItems(3) = ""
      rs.MoveNext:
      numero_items_plantas = numero_items_plantas + 1
    Wend
   rs.Close
   
   rs.Open "select segment1, description from xxvia_system_items_b where organization_id = " + var_unidad_organizacional + " order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
       Set list_item = lv_articulos.ListItems.Add(, , rs!SEGMENT1)
       list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
       list_item.SubItems(2) = " "
       rs.MoveNext
   Wend
   rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_movimientos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_movimientos.selectedItem.Index
      If lv_movimientos.selectedItem.SubItems(2) = "*" Then
         lv_movimientos.selectedItem.SubItems(2) = ""
         lv_movimientos.ListItems.Item(i).Bold = False
         lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_movimientos.Refresh
      Else
         lv_movimientos.selectedItem.SubItems(2) = "*"
         lv_movimientos.ListItems.Item(i).Bold = True
         lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_movimientos.Refresh
      End If
   End If
End Sub

Private Sub lv_plantas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   i = lv_plantas.selectedItem.Index
      If lv_plantas.selectedItem.SubItems(3) = "*" Then
         lv_plantas.selectedItem.SubItems(3) = ""
         lv_plantas.ListItems.Item(i).Bold = False
         lv_plantas.ListItems.Item(i).ForeColor = &H80000012
         lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_plantas.Refresh
      Else
         lv_plantas.selectedItem.SubItems(3) = "*"
         lv_plantas.ListItems.Item(i).Bold = True
         lv_plantas.ListItems.Item(i).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_plantas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_plantas.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_plantas.Refresh
      End If
   End If
End Sub

Private Sub mon_mes1_DateDblClick(ByVal DateDblClicked As Date)
   txt_inicio = mon_mes1.Value
   mon_mes1.Visible = False
End Sub

Private Sub mon_mes1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_inicio = mon_mes1.Value
      mon_mes1.Visible = False
   End If
   If KeyAscii = 27 Then
      mon_mes1.Visible = False
   End If
End Sub

Private Sub mon_mes1_LostFocus()
   mon_mes1.Visible = False
End Sub

Private Sub mon_mes2_DateDblClick(ByVal DateDblClicked As Date)
   txt_fin = mon_mes2.Value
   mon_mes2.Visible = False
End Sub

Private Sub mon_mes2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fin = mon_mes2.Value
      mon_mes2.Visible = False
   End If
   If KeyAscii = 27 Then
      mon_mes2.Visible = False
   End If
End Sub

Private Sub mon_mes2_LostFocus()
   mon_mes2.Visible = False
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_buscar_LostFocus()
   Call pro_busca_registro(lv_articulos, txt_buscar, False)
   txt_buscar = ""
End Sub
