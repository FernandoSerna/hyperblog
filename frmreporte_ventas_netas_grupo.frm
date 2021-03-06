VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_ventas_netas_grupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ventas Netas por Grupo"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   " Canales de Venta "
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   5625
      Begin MSComctlLib.ListView lv_canales 
         Height          =   2025
         Left            =   45
         TabIndex        =   18
         Top             =   210
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3572
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
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Grupos "
      Height          =   3270
      Left            =   120
      TabIndex        =   5
      Top             =   2970
      Width           =   5625
      Begin VB.CheckBox chk_detalle 
         Caption         =   "Detalle"
         Height          =   285
         Left            =   2205
         TabIndex        =   14
         Top             =   255
         Width           =   945
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_ventas_netas_grupo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_ventas_netas_grupo.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_ventas_netas_grupo.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selecci?n Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_ventas_netas_grupo.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_ventas_netas_grupo.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   8
         Top             =   540
         Width           =   5565
      End
      Begin VB.CheckBox chk_comparativo 
         Caption         =   "Comparativo"
         Height          =   285
         Left            =   3435
         TabIndex        =   7
         Top             =   255
         Width           =   1485
      End
      Begin VB.TextBox txt_busqueda 
         Height          =   375
         Left            =   1155
         TabIndex        =   6
         Top             =   735
         Width           =   4395
      End
      Begin MSComctlLib.ListView lv_grupos 
         Height          =   2025
         Left            =   45
         TabIndex        =   15
         Top             =   1170
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3572
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
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda:"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   825
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   4
      Top             =   360
      Width           =   5685
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_ventas_netas_grupo.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Generar Reporte "
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5370
      Picture         =   "frmreporte_ventas_netas_grupo.frx":0B5C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " A?o "
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   6255
      Width           =   5640
      Begin VB.TextBox txt_a?o 
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
         Height          =   690
         Left            =   1560
         TabIndex        =   1
         Top             =   165
         Width           =   1950
      End
   End
End
Attribute VB_Name = "frmreporte_ventas_netas_grupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer
Private Sub cmd_imprimir_Click()
   Dim pError As ADODB.Error
   'On Error GoTo salir:
   Dim var_consecutivo As Double
   Dim var_contador As Double
   Dim var_cadena As String
   Dim var_cadena_2 As String
   Dim var_contador_errores As Integer
   var_contador_errores = 0
   If IsNumeric(txt_a?o) Then
      var_contador = 0
      var_cadena = ""
      var_cadena_2 = ""
      If Me.chk_comparativo = 1 Then
         For var_i = 1 To lv_grupos.ListItems.Count
             lv_grupos.ListItems.Item(var_i).Selected = True
             If Me.chk_detalle.Value = 0 Then
                If lv_grupos.selectedItem.SubItems(2) = "*" Then
                   var_contador = var_contador + 1
                   If Len(Trim(var_cadena)) = 0 Then
                      var_cadena = var_cadena + "{VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena = var_cadena + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                   If Len(Trim(var_cadena_2)) = 0 Then
                      var_cadena_2 = var_cadena_2 + " {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena_2 = var_cadena_2 + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                End If
             Else
                If lv_grupos.selectedItem.SubItems(2) = "*" Then
                   var_contador = var_contador + 1
                   If Len(Trim(var_cadena)) = 0 Then
                      var_cadena = var_cadena + "{VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena = var_cadena + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                   If Len(Trim(var_cadena_2)) = 0 Then
                      var_cadena_2 = var_cadena_2 + " {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena_2 = var_cadena_2 + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                End If
             End If
         Next var_i
         If var_contador > 0 Then
             
            Frmmenu2.StatusBar1.Panels(1).Text = "Procesando informaci?n, espere un momento"
            cnn.CommandTimeout = 360
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_VENTAS_NETAS_CLIENTES", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_VENTAS_NETAS_CLIENTES (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
         
            var_n = Me.lv_grupos.ListItems.Count
            For var_i = 1 To var_n
                lv_grupos.ListItems.Item(var_i).Selected = True
                If lv_grupos.selectedItem.SubItems(2) = "*" Then
                   rs.Open "insert into TB_TEMP_REPORTE_VENTAS_NETAS_GRUPOS (inte_tem_consecutivo, VCHA_GAC_GRUPO_aCTUAL_ID) values (" + CStr(var_consecutivo) + ",'" + lv_grupos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            rs.Open "EXEC SP_REPORTE_VENTAS_NETAS_CLIENTES_TITULAR_2 " + CStr(var_consecutivo) + ", " + txt_a?o + ", '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            If Me.chk_detalle.Value = 0 Then
               
               VAR_CADENA_canal = ""
               For var_i = 1 To Me.lv_canales.ListItems.Count
                   lv_canales.ListItems.Item(var_i).Selected = True
                   If lv_canales.selectedItem.SubItems(2) = "*" Then
                      If Len(Trim(VAR_CADENA_canal)) = 0 Then
                         VAR_CADENA_canal = " AND ({VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      Else
                         VAR_CADENA_canal = VAR_CADENA_canal + " OR {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      End If
                   End If
               Next var_i
               VAR_CADENA_canal = VAR_CADENA_canal + ")"
               
               
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_grupo_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and (" + var_cadena + ")" + VAR_CADENA_canal
            Else
               
               VAR_CADENA_canal = ""
               For var_i = 1 To Me.lv_canales.ListItems.Count
                   lv_canales.ListItems.Item(var_i).Selected = True
                   If lv_canales.selectedItem.SubItems(2) = "*" Then
                      If Len(Trim(VAR_CADENA_canal)) = 0 Then
                         VAR_CADENA_canal = " AND ({VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      Else
                         VAR_CADENA_canal = VAR_CADENA_canal + " OR {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      End If
                   End If
               Next var_i
               VAR_CADENA_canal = VAR_CADENA_canal + ")"
               
               
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_grupo_detalle.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_VENTAS_NETAS_GRUPOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and (" + var_cadena + ")" + VAR_CADENA_canal
            End If
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            If Me.chk_detalle.Value = 0 Then
               archivo = "c:\reportessid\Reporte_ventas_grupos_concentrado" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            Else
               archivo = "c:\reportessid\Reporte_ventas_grupos_detalle" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            End If
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CLIENTES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_GRUPOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Frmmenu2.StatusBar1.Panels(1).Text = ""
         Else
            MsgBox "No se a seleccionado un grupo", vbOKOnly, "ATENCION"
         End If
      Else
         For var_i = 1 To lv_grupos.ListItems.Count
             lv_grupos.ListItems.Item(var_i).Selected = True
             If Me.chk_detalle.Value = 0 Then
                If lv_grupos.selectedItem.SubItems(2) = "*" Then
                   var_contador = var_contador + 1
                   If Len(Trim(var_cadena)) = 0 Then
                      var_cadena = var_cadena + "{VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena = var_cadena + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                   If Len(Trim(var_cadena_2)) = 0 Then
                      var_cadena_2 = var_cadena_2 + " {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena_2 = var_cadena_2 + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                End If
             Else
                If lv_grupos.selectedItem.SubItems(2) = "*" Then
                   var_contador = var_contador + 1
                   If Len(Trim(var_cadena)) = 0 Then
                      var_cadena = var_cadena + "{VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena = var_cadena + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                   If Len(Trim(var_cadena_2)) = 0 Then
                      var_cadena_2 = var_cadena_2 + " {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   Else
                      var_cadena_2 = var_cadena_2 + " or {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_GAC_GRUPO_ACTUAL_ID} = '" + lv_grupos.selectedItem + "'"
                   End If
                End If
             End If
         Next var_i
         If var_contador > 0 Then
            
            VAR_CADENA_canal = ""
            For var_i = 1 To Me.lv_canales.ListItems.Count
                lv_canales.ListItems.Item(var_i).Selected = True
                If lv_canales.selectedItem.SubItems(2) = "*" Then
                   If Len(Trim(VAR_CADENA_canal)) = 0 Then
                      VAR_CADENA_canal = " AND ({VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                   Else
                      VAR_CADENA_canal = VAR_CADENA_canal + " OR {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                   End If
                End If
            Next var_i
            VAR_CADENA_canal = VAR_CADENA_canal + ")"
            
            Frmmenu2.StatusBar1.Panels(1).Text = "Procesando informaci?n, espere un momento"
            cnn.CommandTimeout = 360
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_VENTAS_NETAS_CLIENTES", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_VENTAS_NETAS_CLIENTES (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
         
            var_n = Me.lv_grupos.ListItems.Count
            For var_i = 1 To var_n
                lv_grupos.ListItems.Item(var_i).Selected = True
                If lv_grupos.selectedItem.SubItems(2) = "*" Then
                   rs.Open "insert into TB_TEMP_REPORTE_VENTAS_NETAS_GRUPOS (inte_tem_consecutivo, VCHA_GAC_GRUPO_aCTUAL_ID) values (" + CStr(var_consecutivo) + ",'" + lv_grupos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_i
            rs.Open "EXEC SP_REPORTE_VENTAS_NETAS_CLIENTES_TITULAR_2 " + CStr(var_consecutivo) + ", " + txt_a?o + ", '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            If Me.chk_detalle.Value = 0 Then
               
               VAR_CADENA_canal = ""
               For var_i = 1 To Me.lv_canales.ListItems.Count
                   lv_canales.ListItems.Item(var_i).Selected = True
                   If lv_canales.selectedItem.SubItems(2) = "*" Then
                      If Len(Trim(VAR_CADENA_canal)) = 0 Then
                         VAR_CADENA_canal = " AND ({VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      Else
                         VAR_CADENA_canal = VAR_CADENA_canal + " OR {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      End If
                   End If
               Next var_i
               VAR_CADENA_canal = VAR_CADENA_canal + ")"
               
               
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_GRUPO_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_VENTAS_NETAS_GRUPOS_CONCENTRADO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and (" + var_cadena + ")" + VAR_CADENA_canal
            Else
               
               VAR_CADENA_canal = ""
               For var_i = 1 To Me.lv_canales.ListItems.Count
                   lv_canales.ListItems.Item(var_i).Selected = True
                   If lv_canales.selectedItem.SubItems(2) = "*" Then
                      If Len(Trim(VAR_CADENA_canal)) = 0 Then
                         VAR_CADENA_canal = " AND ({VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      Else
                         VAR_CADENA_canal = VAR_CADENA_canal + " OR {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_CAN_CANAL_VENTA_ID} = '" + lv_canales.selectedItem + "'"
                      End If
                   End If
               Next var_i
               VAR_CADENA_canal = VAR_CADENA_canal + ")"
               
               
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_GRUPO_DETALLE.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_VENTAS_NETAS_GRUPOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_VENTAS_NETAS_GRUPOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and (" + var_cadena + ")" + VAR_CADENA_canal
            End If
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            If Me.chk_detalle.Value = 0 Then
               archivo = "c:\reportessid\Reporte_ventas_grupos_concentrado" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            Else
               archivo = "c:\reportessid\Reporte_ventas_grupos_detalle" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            End If
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CLIENTES where inte_tem_consecutivo = " + CStr(var_consecutivo)
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_GRUPOS where inte_tem_consecutivo = " + CStr(var_consecutivo)
            Frmmenu2.StatusBar1.Panels(1).Text = ""
         Else
            MsgBox "No se a seleccionado un grupo", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "A?o incorrecto", vbOKOnly, "ATENCION"
      txt_a?o = ""
   End If
   Exit Sub
salir:
   If Err.Number = -2147217871 Then
      var_si = MsgBox("El sistema a marcado tiempo de espera agotado, ?Desea continuar?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Resume
         var_contador_errores = var_contador_errores + 1
         If var_contador_errores = 4 Then
            MsgBox "A surgido un error al conectarce a la base de datos", vbOKOnly, "ATENCION"
            Exit Sub
         End If
      Else
         Exit Sub
      End If
  
   Else
      MsgBox "A surgido un error", vbOKOnly, "ATENCION"
   End If
    
End Sub

Private Sub cmd_invertir_Click()
   n = lv_grupos.ListItems.Count
   For i = 1 To n
      lv_grupos.ListItems.Item(i).Selected = True
      If lv_grupos.selectedItem.SubItems(2) = "*" Then
         lv_grupos.selectedItem.SubItems(2) = ""
         lv_grupos.ListItems.Item(i).Bold = False
         lv_grupos.ListItems.Item(i).ForeColor = &H80000012
         lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_grupos.selectedItem.SubItems(2) = "*"
         lv_grupos.ListItems.Item(i).Bold = True
         lv_grupos.ListItems.Item(i).ForeColor = &HFF0000
         lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_grupos.selectedItem.Index
   If lv_grupos.selectedItem.SubItems(2) = "*" Then
      lv_grupos.selectedItem.SubItems(2) = ""
      lv_grupos.ListItems.Item(i).Bold = False
      lv_grupos.ListItems.Item(i).ForeColor = &H80000012
      lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_grupos.Refresh
   Else
      lv_grupos.selectedItem.SubItems(2) = "*"
      lv_grupos.ListItems.Item(i).Bold = True
      lv_grupos.ListItems.Item(i).ForeColor = &HFF0000
      lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_grupos.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_grupos.ListItems.Count
   For i = 1 To n
      lv_grupos.ListItems.Item(i).Selected = True
      lv_grupos.selectedItem.SubItems(2) = ""
      lv_grupos.ListItems.Item(i).Bold = False
      lv_grupos.ListItems.Item(i).ForeColor = &H80000012
      lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_grupos.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_grupos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_grupos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_grupos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_grupos.selectedItem.SubItems(2) = "*"
         lv_grupos.ListItems.Item(i).Bold = True
         lv_grupos.ListItems.Item(i).ForeColor = &HFF0000
         lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_grupos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_grupos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_grupos.ListItems.Count
   For i = 1 To n
      lv_grupos.ListItems.Item(i).Selected = True
      lv_grupos.selectedItem.SubItems(2) = "*"
      lv_grupos.ListItems.Item(i).Bold = True
      lv_grupos.ListItems.Item(i).ForeColor = &HFF0000
      lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_grupos.Refresh
End Sub

Private Sub Form_Load()
Dim dl As Long                                 ' Valor devuelto por la funci?n API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripci?n del DSN
Dim sDsnName As String                  ' Nombre del DSN

   cnn.Close
   cnn.Open var_conexion_string_distribucion

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se crear? un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminar? un DSN de sistema
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
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   
   
   
   Me.txt_a?o = Year(Date)
   var_cadena_seguridad = ""
   Top = 0
   Left = 3200
   txt_inicio = Date
   txt_fin = Date
   'opt_linea = True
   
   rs.Open "select distinct vcha_Can_canal_venta_id, vcha_can_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rut_ruta_id is not null and vcha_tit_nombre <> '' and inte_can_reporte= 1 order by vcha_can_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_canales.ListItems.Add(, , rs!vcha_can_canal_venta_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
   If lv_canales.ListItems.Count > 7 Then
      lv_canales.ColumnHeaders(2).Width = 4220
   Else
      lv_canales.ColumnHeaders(2).Width = 4499.71
   End If
   
   
'   rs.Open "select distinct vcha_gac_grupo_Actual_id, vcha_gac_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_actual_id is not null and vcha_tit_nombre <> '' order by vcha_gac_nombre ", cnn, adOpenDynamic, adLockOptimistic
'   numero_items_ALMACENES = 0
'   While Not rs.EOF
'      Set list_item = lv_grupos.ListItems.Add(, , rs!vcha_gac_grupo_Actual_id)
'      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_gac_NOMBRE), "", rs!VCHA_gac_NOMBRE)
'      list_item.SubItems(2) = ""
'      rs.MoveNext:
'   Wend
'   rs.Close
'   If lv_grupos.ListItems.Count > 7 Then
'      lv_grupos.ColumnHeaders(2).Width = 4220
'   Else
'      lv_grupos.ColumnHeaders(2).Width = 4499.71
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_salidas)
End Sub

Private Sub lv_canales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_canales, ColumnHeader)
End Sub

Private Sub lv_canales_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_canales.selectedItem.Index
      If lv_canales.selectedItem.SubItems(2) = "*" Then
         lv_canales.selectedItem.SubItems(2) = ""
         lv_canales.ListItems.Item(i).Bold = False
         lv_canales.ListItems.Item(i).ForeColor = &H80000012
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_canales.Refresh
      Else
         lv_canales.selectedItem.SubItems(2) = "*"
         lv_canales.ListItems.Item(i).Bold = True
         lv_canales.ListItems.Item(i).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_canales.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_canales.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_canales.Refresh
      End If
      var_cadena = ""
      For var_i = 1 To lv_canales.ListItems.Count
          lv_canales.ListItems.Item(var_i).Selected = True
          If lv_canales.selectedItem.SubItems(2) = "*" Then
              If var_cadena = "" Then
                 var_cadena = "(vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'"
              Else
                 var_cadena = var_cadena + " or vcha_can_canal_venta_id = '" + lv_canales.selectedItem + "'"
              End If
          End If
      Next var_i
      lv_grupos.ListItems.Clear
      If var_cadena <> "" Then
         rs.Open "select distinct vcha_gac_grupo_Actual_id, vcha_gac_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_actual_id is not null and vcha_tit_nombre <> '' and " + var_cadena + ") order by vcha_gac_nombre ", cnn, adOpenDynamic, adLockOptimistic
         numero_items_ALMACENES = 0
         While Not rs.EOF
             Set list_item = lv_grupos.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
             list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
             list_item.SubItems(2) = ""
             rs.MoveNext:
         Wend
         rs.Close
         If lv_grupos.ListItems.Count > 7 Then
            lv_grupos.ColumnHeaders(2).Width = 4220
         Else
            lv_grupos.ColumnHeaders(2).Width = 4499.71
         End If
      End If
   End If
End Sub

Private Sub lv_grupos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_grupos, ColumnHeader)
End Sub

Private Sub lv_grupos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_grupos.ListItems.Count > 0 Then
         i = lv_grupos.selectedItem.Index
         If lv_grupos.selectedItem.SubItems(2) = "*" Then
            lv_grupos.selectedItem.SubItems(2) = ""
            lv_grupos.ListItems.Item(i).Bold = False
            lv_grupos.ListItems.Item(i).ForeColor = &H80000012
            lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_grupos.Refresh
         Else
            lv_grupos.selectedItem.SubItems(2) = "*"
            lv_grupos.ListItems.Item(i).Bold = True
            lv_grupos.ListItems.Item(i).ForeColor = &HFF0000
            lv_grupos.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_grupos.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_grupos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_grupos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_grupos.Refresh
         End If
      End If
   End If
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select top 1  vcha_gac_nombre from vw_clientes where vcha_gac_nombre like '%" + Me.txt_busqueda + "%' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(lv_grupos, rs!vcha_gac_nombre, False)
      Else
         MsgBox "No existen coincidencias", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

