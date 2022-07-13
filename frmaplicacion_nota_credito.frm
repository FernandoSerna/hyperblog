VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaplicacion_nota_credito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicación de Notas de Crédito"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   825
      Picture         =   "frmaplicacion_nota_credito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Ver notas de crédito sin aplicar"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_importe_eliminar 
      Height          =   915
      Left            =   2940
      TabIndex        =   24
      Top             =   3225
      Width           =   2565
      Begin VB.TextBox txt_importe_eliminar 
         Height          =   360
         Left            =   60
         TabIndex        =   25
         Top             =   420
         Width           =   2400
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Importe a eliminar"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   15
         TabIndex        =   26
         Top             =   15
         Width           =   2535
      End
   End
   Begin VB.Frame frm_importe_aplicar 
      Height          =   915
      Left            =   2940
      TabIndex        =   21
      Top             =   3210
      Width           =   2565
      Begin VB.TextBox txt_importe_aplicar 
         Height          =   360
         Left            =   60
         TabIndex        =   22
         Top             =   420
         Width           =   2400
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Importe a aplicar"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   15
         TabIndex        =   23
         Top             =   15
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Documentos a Aplicar "
      Height          =   3765
      Left            =   120
      TabIndex        =   18
      Top             =   1995
      Width           =   7425
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   3090
         Left            =   60
         TabIndex        =   11
         Top             =   255
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5450
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
            Text            =   "Documento"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Serie"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Número"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Aplicar"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label lbl_importe 
         Alignment       =   1  'Right Justify
         Caption         =   "999,999,999.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   3390
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4290
         TabIndex        =   19
         Top             =   3375
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7245
      Picture         =   "frmaplicacion_nota_credito.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command9 
      Height          =   315
      Left            =   165
      Picture         =   "frmaplicacion_nota_credito.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Desmarcar Todos Alt + D"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   495
      Picture         =   "frmaplicacion_nota_credito.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   45
      TabIndex        =   17
      Top             =   345
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la Nota de Crédito"
      Height          =   1545
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   7425
      Begin VB.TextBox txt_tipo 
         Height          =   350
         Left            =   4575
         TabIndex        =   6
         Top             =   315
         Width           =   480
      End
      Begin VB.TextBox txt_serie 
         Height          =   350
         Left            =   885
         TabIndex        =   4
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txt_importe 
         Height          =   350
         Left            =   885
         TabIndex        =   10
         Top             =   1065
         Width           =   1605
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   350
         Left            =   2115
         TabIndex        =   9
         Top             =   690
         Width           =   5190
      End
      Begin VB.TextBox txt_cliente 
         Height          =   350
         Left            =   885
         TabIndex        =   8
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox txt_fecha 
         Height          =   350
         Left            =   5715
         TabIndex        =   7
         Top             =   315
         Width           =   1605
      End
      Begin VB.TextBox txt_numero 
         Height          =   350
         Left            =   2820
         TabIndex        =   5
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   4140
         TabIndex        =   28
         Top             =   390
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   393
         Width           =   405
      End
      Begin VB.Label lbl_moneda 
         Height          =   345
         Left            =   2805
         TabIndex        =   16
         Top             =   1050
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   1140
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   5130
         TabIndex        =   14
         Top             =   393
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   765
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   2205
         TabIndex        =   12
         Top             =   390
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmaplicacion_nota_credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_aceptar_pedidos_Click()
   If CDbl(Me.lbl_importe) = CDbl(Me.txt_importe) Then
      var_si = MsgBox("¿Desea aplicar la nota crédito", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la aplicación de la nota de crédito", vbYesNo, "ATENCION")
         If var_si = 6 Then
            For var_j = 1 To lv_facturas.ListItems.Count
                lv_facturas.ListItems.Item(var_j).Selected = True
                If CDbl(Me.lv_facturas.selectedItem.SubItems(5)) > 0 Then
                   var_cadena = "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_cargo, inte_ecu_numero_cargo,floa_ecu_importe_cargo, vcha_ecu_movimiento_abono, vcha_ecu_serie_abono, inte_ecu_numero_abono, floa_ecu_importe_abono) values"
                   var_cadena = var_cadena + " ('" + var_empresa + "', '" + lv_facturas.selectedItem + "','" + lv_facturas.selectedItem.SubItems(1) + "'," + Me.lv_facturas.selectedItem.SubItems(2) + ",0,'" + Me.txt_tipo + "','" + Me.txt_serie + "'," + Me.txt_numero + "," + Me.lv_facturas.selectedItem.SubItems(5) + ")"
                   rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                   cnn.BeginTrans
                   rs.Open "UPDATE TB_ENCABEZADO_CARTERA SET INTE_CAR_NOTA_CREDITO_APLICADA = 1, DTIM_CAR_FECHA = GETDATE() WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'NC' AND VCHA_CAR_DOCUMENTO = '" + Me.txt_tipo + "' AND VCHA_SER_sERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                   'rs.Open "update tb_saldo set floa_Sal_importe = floa_sal_importe - "+
                   cnn.CommitTrans
                End If
            Next var_j
         End If
      End If
      MsgBox "Se termino la aplicacion de la nota de crédito", vbOKOnly, "ATENCIO"
      Me.txt_cliente = ""
      Me.txt_fecha = ""
      Me.txt_importe = ""
      Me.txt_importe_aplicar = ""
      Me.txt_importe_eliminar = ""
      Me.txt_nombre_cliente = ""
      Me.txt_numero = ""
      Me.txt_serie = ""
      Me.lbl_importe = "0.00"
      Me.lbl_moneda = ""
      Me.txt_tipo = ""
      Me.lv_facturas.ListItems.Clear
   Else
      MsgBox "El importe a aplicar debe de ser igual que el importe de la nota de crédito", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
         Set reporte = appl.OpenReport(App.Path + "\rep_notas_credito_sin_aplicar.rpt")
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de notas de crédito sin aplicar"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_notas_credito_sin_aplicar.rpt")
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\notas_credito_sin_aplicar" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
         
         
         
         
End Sub

Private Sub Command9_Click()
   Me.txt_cliente = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe_aplicar = ""
   Me.txt_importe_eliminar = ""
   Me.txt_nombre_cliente = ""
   Me.txt_numero = ""
   Me.txt_serie = ""
   Me.lbl_importe = "0.00"
   Me.lbl_moneda = ""
   Me.txt_tipo = ""
   Me.lv_facturas.ListItems.Clear
   Me.txt_serie.SetFocus
End Sub

Private Sub Form_Load()
   Top = 800
   Left = 2000
   Me.frm_importe_aplicar.Visible = False
   Me.frm_importe_eliminar.Visible = False
   Me.lbl_importe = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_facturas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      Me.frm_importe_aplicar.Visible = True
      Me.txt_importe_aplicar = ""
      Me.txt_importe_aplicar.SetFocus
   End If
   If KeyCode = 114 Then
      Me.frm_importe_eliminar.Visible = True
      Me.txt_importe_eliminar = ""
      Me.txt_importe_eliminar.SetFocus
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_importe_aplicar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_importe_aplicar) Then
         If CDbl(Me.lv_facturas.selectedItem.SubItems(4)) >= CDbl(Me.txt_importe_aplicar) Then
            If CDbl(Me.txt_importe_aplicar) + CDbl(Me.lbl_importe) <= CDbl(Me.txt_importe) Then
               Me.lv_facturas.selectedItem.SubItems(5) = CDbl(Me.lv_facturas.selectedItem.SubItems(5)) + Me.txt_importe_aplicar
               Me.lbl_importe = CDbl(Me.lbl_importe) + CDbl(Me.txt_importe_aplicar)
            Else
               MsgBox "El importe a aplicar no puede superar al importe de la nota de crédito", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El importe no puede ser mayor al importe de la factura", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
      End If
      If lv_facturas.ListItems.Count > 0 Then
         Me.lv_facturas.SetFocus
      Else
         Me.frm_importe_aplicar.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_importe_aplicar.Visible = False
   End If
End Sub

Private Sub txt_importe_aplicar_LostFocus()
   Me.frm_importe_aplicar.Visible = False
End Sub

Private Sub txt_importe_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_importe_eliminar) Then
         If CDbl(Me.txt_importe_eliminar) <= lv_facturas.selectedItem.SubItems(5) Then
            Me.lv_facturas.selectedItem.SubItems(5) = CDbl(Me.lv_facturas.selectedItem.SubItems(5)) - CDbl(Me.txt_importe_eliminar)
            Me.lbl_importe = CDbl(Me.lbl_importe) - CDbl(Me.txt_importe_eliminar)
         Else
            MsgBox "Importe a eliminar supera al importe por aplicar", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Importe a eliminar incorrecto", vbOKOnly, "ATENCION"
      End If
      If lv_facturas.ListItems.Count > 0 Then
         Me.lv_facturas.SetFocus
      Else
        Me.frm_importe_eliminar.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_importe_eliminar.Visible = False
   End If
End Sub

Private Sub txt_importe_eliminar_LostFocus()
   Me.frm_importe_eliminar.Visible = False
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   If Trim(Me.txt_numero) <> "" Then
      If IsNumeric(Me.txt_numero) Then
         rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + Me.txt_serie + "' and inte_car_numero = " + Me.txt_numero + " and vcha_car_tipo_documento = 'NC' AND INTE_CAR_NOTA_CREDITO_APLICADA = 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_tipo = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
            Me.txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            Me.txt_importe = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)
            Me.txt_fecha = Format(IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha), "Short Date")
            rsaux2.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               Me.txt_nombre_cliente = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
            End If
            rsaux2.Close
            rsaux2.Open "select * from vw_saldos_facturas where vcha_cli_clave_id = '" + Me.txt_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  Set list_item = Me.lv_facturas.ListItems.Add(, , rsaux2!vcha_Car_documento)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2!VCHA_SER_SERIE_ID), "", rsaux2!VCHA_SER_SERIE_ID)
                  list_item.SubItems(2) = IIf(IsNull(rsaux2!inte_Car_numero), "", rsaux2!inte_Car_numero)
                  list_item.SubItems(3) = Format(IIf(IsNull(rsaux2!dtim_Car_fecha), "", rsaux2!dtim_Car_fecha), "Short Date")
                  list_item.SubItems(4) = Format(IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), "", rsaux2!FLOA_sAL_IMPORTE), "###,###,##0.00")
                  list_item.SubItems(5) = 0
                  rsaux2.MoveNext
            Wend
            Me.lbl_importe = "0.00"
            
            
            rsaux2.Close
         Else
            MsgBox "No existen notas de crédito por aplicar", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de nota de crédito incorrecta", vbOKOnly, "ATENCION"
         Me.lv_facturas.ListItems.Clear
         Me.txt_cliente = ""
         Me.txt_fecha = ""
         Me.txt_nombre_cliente = ""
         Me.txt_importe = ""
         Me.lbl_importe = ""
         Me.lbl_moneda = ""
      End If
   Else
      MsgBox "Debe de indicar el número de la nota de crédito", vbOKOnly, "ATENCION"
      Me.lv_facturas.ListItems.Clear
      Me.txt_cliente = ""
      Me.txt_fecha = ""
      Me.txt_nombre_cliente = ""
      Me.txt_importe = ""
      Me.lbl_importe = ""
      Me.lbl_moneda = ""
      
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
