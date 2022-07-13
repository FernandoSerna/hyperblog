VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmver_pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Height          =   375
      Left            =   90
      Picture         =   "frmver_pedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   30
      Width           =   375
   End
   Begin VB.CommandButton cmd_salir 
      Height          =   375
      Left            =   9840
      Picture         =   "frmver_pedido.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   30
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   75
      TabIndex        =   11
      Top             =   435
      Width           =   10185
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detalle del pedido "
      Height          =   4290
      Left            =   120
      TabIndex        =   2
      Top             =   2430
      Width           =   10095
      Begin MSComctlLib.ListView lv_pedido 
         Height          =   3645
         Left            =   75
         TabIndex        =   15
         Top             =   210
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   6429
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Pedido"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Orden de Surtido"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Faltan"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_cantidad_faltante 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8460
         TabIndex        =   19
         Top             =   3870
         Width           =   1320
      End
      Begin VB.Label lbl_cantidad_surtir 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6825
         TabIndex        =   18
         Top             =   3870
         Width           =   1320
      End
      Begin VB.Label lbl_cantidad_pedida 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5475
         TabIndex        =   17
         Top             =   3870
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   360
         Left            =   4365
         TabIndex        =   16
         Top             =   3862
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Pedido "
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   525
      Width           =   10080
      Begin VB.TextBox txt_fecha 
         Height          =   350
         Left            =   4020
         TabIndex        =   14
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   350
         Left            =   2400
         TabIndex        =   10
         Top             =   1155
         Width           =   7530
      End
      Begin VB.TextBox txt_cliente 
         Height          =   350
         Left            =   825
         TabIndex        =   9
         Top             =   1155
         Width           =   1545
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   350
         Left            =   2400
         TabIndex        =   7
         Top             =   780
         Width           =   7530
      End
      Begin VB.TextBox txt_agente 
         Height          =   350
         Left            =   825
         TabIndex        =   6
         Top             =   780
         Width           =   1545
      End
      Begin VB.TextBox txt_pedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   825
         TabIndex        =   1
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1230
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   855
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3345
         TabIndex        =   4
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmver_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If Me.txt_pedido <> "" Then
      If IsNumeric(Me.txt_pedido) Then
         If Trim(txt_pedido) <> "" Then
            rs.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_estatus = IIf(IsNull(rs!char_ped_estatus), "", rs!char_ped_estatus)
               If Trim(var_estatus) <> "" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_PEDIDos_vt.rpt")
                  reporte.RecordSelectionFormula = "{VW_PEDIDOS.INTE_PED_NUMERO} = " + txt_pedido
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Pedidos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_PEDIDos_vt.rpt")
                     reporte.RecordSelectionFormula = "{VW_PEDIDOS.INTE_PED_NUMERO} = " + txt_pedido
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Pedido_" + Me.txt_pedido + "_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               Else
                  MsgBox "El pedido aun no a sido cerrado", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 300
   Left = 700
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_pedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pedido, ColumnHeader)
End Sub

Private Sub lv_pedido_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call pro_enfoque(KeyAscii)
  Else
     If KeyAscii = 27 Then
        Unload Me
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If keysacii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_pedido_Change()
   Me.txt_agente = ""
   Me.txt_cliente = ""
   Me.txt_nombre_agente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_fecha = ""
   Me.lbl_cantidad_faltante = "0.00"
   Me.lbl_cantidad_pedida = "0.00"
   Me.lbl_cantidad_surtir = "0.00"
   Me.lv_pedido.ListItems.Clear
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_pedido <> "" Then
         If IsNumeric(Me.txt_pedido) Then
            If Trim(txt_pedido) <> "" Then
               Me.txt_agente = ""
               Me.txt_nombre_agente = ""
               Me.txt_cliente = ""
               Me.txt_nombre_cliente = ""
               Me.txt_fecha = ""
               Me.lbl_cantidad_faltante = "0.00"
               Me.lbl_cantidad_pedida = "0.00"
               Me.lbl_cantidad_surtir = "0.00"
               Me.lv_pedido.ListItems.Clear
               rs.Open "select isnull(inte_ped_autorizo,0) as inte_ped_autorizo, isnull(char_ped_estatus,'') as char_ped_estatus, vcha_Cli_clave_id, vcha_age_agente_id, dtim_ped_fecha from tb_encabezado_pedidos where inte_ped_numero = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Dim list_item As ListItem
                  var_estatus = IIf(IsNull(rs!char_ped_estatus), "", rs!char_ped_estatus)
                  If Trim(var_estatus) <> "" Then
                     rsaux.Open "select vcha_age_agente_id, vcha_Age_nombre from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Me.txt_agente = IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID)
                        Me.txt_nombre_agente = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
                     Else
                        Me.txt_agente = ""
                        Me.txt_nombre_agente = ""
                     End If
                     rsaux.Close
                     Me.txt_fecha = Format(rs!dtim_ped_fecha, "Short Date")
                     rsaux.Open "select * from tb_clientes where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Me.txt_cliente = IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id)
                        Me.txt_nombre_cliente = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
                     Else
                        Me.txt_cliente = ""
                        Me.txt_nombre_cliente = ""
                     End If
                     rsaux.Close
                     var_cadena = "SELECT SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AS FLOA_ORS_CANTIDAD_SURTIR, dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD , dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_DET_ORDEN_SURTIDO INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON"
                     var_cadena = var_cadena + " dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO AND dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_DETALLE_PEDIDOS.VCHA_ART_ARTICULO_ID INNER JOIN"
                     var_cadena = var_cadena + " dbo.TB_ARTICULOS ON dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID GROUP BY dbo.TB_DET_ORDEN_SURTIDO.VCHA_ART_ARTICULO_ID, dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD , dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL Having (dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = " + Me.txt_pedido + ")"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux.EOF
                           Set list_item = Me.lv_pedido.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_español), "", rsaux!vcha_art_nombre_español)
                           list_item.SubItems(2) = Format(IIf(IsNull(rsaux!FLOA_PED_CANTIDAD), 0, rsaux!FLOA_PED_CANTIDAD), "###,###,##0.00")
                           Me.lbl_cantidad_pedida = Format(CDbl(Me.lbl_cantidad_pedida) + IIf(IsNull(rsaux!FLOA_PED_CANTIDAD), 0, rsaux!FLOA_PED_CANTIDAD), "###,###,##0.00")
                           list_item.SubItems(3) = Format(IIf(IsNull(rsaux!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                           Me.lbl_cantidad_surtir = Format(CDbl(Me.lbl_cantidad_surtir) + IIf(IsNull(rsaux!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                           list_item.SubItems(4) = Format(CDbl(IIf(IsNull(rsaux!FLOA_PED_CANTIDAD), 0, rsaux!FLOA_PED_CANTIDAD)) - CDbl(IIf(IsNull(rsaux!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux!FLOA_ORS_CANTIDAD_SURTIR)), "###,###,##0.00")
                           Me.lbl_cantidad_faltante = Format(CDbl(Me.lbl_cantidad_faltante) + CDbl(IIf(IsNull(rsaux!FLOA_PED_CANTIDAD), 0, rsaux!FLOA_PED_CANTIDAD)) - CDbl(IIf(IsNull(rsaux!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux!FLOA_ORS_CANTIDAD_SURTIR)), "###,###,##0.00")
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                  Else
                     MsgBox "El pedido aun no a sido cerrado", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
   Call pro_enfoque(KeyAscii)
End Sub
