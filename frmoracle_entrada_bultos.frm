VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_entrada_bultos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de bultos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista_bultos 
      Height          =   2400
      Left            =   1860
      TabIndex        =   32
      Top             =   1500
      Width           =   5970
      Begin MSComctlLib.ListView lv_lista_bultos 
         Height          =   1950
         Left            =   45
         TabIndex        =   33
         Top             =   405
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3440
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   9701
         EndProperty
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Caption         =   " Tipo de bultos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   34
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.Frame frm_lista_rutas 
      Height          =   3015
      Left            =   1755
      TabIndex        =   29
      Top             =   1110
      Width           =   6750
      Begin VB.TextBox txt_ruta_filtrar 
         Height          =   480
         Left            =   75
         TabIndex        =   37
         Top             =   435
         Width           =   6525
      End
      Begin MSComctlLib.ListView lv_lista_rutas 
         Height          =   1950
         Left            =   45
         TabIndex        =   30
         Top             =   960
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   3440
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
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8643
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         Caption         =   " Rutas"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   31
         Top             =   120
         Width           =   6660
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   975
      Left            =   390
      TabIndex        =   26
      Top             =   240
      Width           =   2025
      Begin VB.TextBox txt_busqueda 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   495
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Busqueda de movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   6
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   1950
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8670
      Picture         =   "frmoracle_entrada_bultos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_entrada_bultos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   705
      Picture         =   "frmoracle_entrada_bultos.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmoracle_entrada_bultos.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   15
      TabIndex        =   24
      Top             =   255
      Width           =   8970
   End
   Begin VB.TextBox txt_total 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6525
      TabIndex        =   19
      Top             =   6870
      Width           =   2430
   End
   Begin VB.Frame Frame4 
      Height          =   795
      Left            =   60
      TabIndex        =   11
      Top             =   2325
      Width           =   8940
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7710
         TabIndex        =   7
         Top             =   195
         Width           =   1125
      End
      Begin VB.TextBox txt_tipo_bulto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1425
         TabIndex        =   6
         Top             =   195
         Width           =   4575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo bulto:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
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
         Left            =   6495
         TabIndex        =   12
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3705
      Left            =   60
      TabIndex        =   10
      Top             =   3105
      Width           =   8940
      Begin MSComctlLib.ListView lv_bultos 
         Height          =   3495
         Left            =   60
         TabIndex        =   25
         Top             =   135
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6165
         View            =   3
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
            Text            =   "Tipo bulto"
            Object.Width           =   12524
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   1275
      Width           =   8940
      Begin VB.TextBox txt_nombre_destino 
         Height          =   360
         Left            =   3090
         TabIndex        =   15
         Top             =   180
         Width           =   5775
      End
      Begin VB.TextBox txt_clave_destino 
         Height          =   360
         Left            =   1395
         TabIndex        =   14
         Top             =   180
         Width           =   1680
      End
      Begin VB.TextBox txt_clave_origen 
         Height          =   360
         Left            =   1395
         TabIndex        =   4
         Top             =   600
         Width           =   1680
      End
      Begin VB.TextBox txt_nombre_origen 
         Height          =   360
         Left            =   3090
         TabIndex        =   5
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
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
         Left            =   195
         TabIndex        =   17
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
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
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   210
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   930
      Left            =   45
      TabIndex        =   8
      Top             =   345
      Width           =   8940
      Begin VB.TextBox txt_estatus 
         Height          =   300
         Left            =   5730
         TabIndex        =   35
         Top             =   150
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt_fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7185
         TabIndex        =   23
         Top             =   285
         Width           =   1605
      End
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
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
         Height          =   465
         Left            =   1380
         TabIndex        =   21
         Top             =   315
         Width           =   1605
      End
      Begin VB.Label lbl_estatus 
         AutoSize        =   -1  'True
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
         Left            =   3765
         TabIndex        =   36
         Top             =   390
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Left            =   6150
         TabIndex        =   22
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
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
         Left            =   165
         TabIndex        =   20
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Total:"
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
      Left            =   5790
      TabIndex        =   18
      Top             =   6900
      Width           =   810
   End
End
Attribute VB_Name = "frmoracle_entrada_bultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_buscar_Click()
   Me.txt_busqueda = ""
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   If Me.txt_estatus = "" Then
      If IsNumeric(Me.txt_folio) Then
         var_si = MsgBox("¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux6.Open "SELECT * FROM TB_ORACLE_CONTROL_BULTOS WHERE NUMERO = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux6.EOF
                  var_cadena = "call XXVIA_PK_CONTROL_COSTALES.xxvia_sp_insert_control_costal (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = var_cadena
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, 88)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(var_unidad_organizacional))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(var_unidad_organizacional))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 40, Me.txt_clave_origen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 40, Me.txt_clave_destino)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, "RC")
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 40, rsaux6!TIPO_BULTO)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux6!NUMERO))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  rsaux6.MoveNext
            Wend
            rsaux6.Close
            rsaux6.Open "UPDATE TB_ORACLE_CONTROL_BULTOS SET ESTATUS = 'I' WHERE NUMERO = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
            Me.txt_estatus = "I"
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_control_bultos.rpt")
            reporte.RecordSelectionFormula = "{TB_ORACLE_CONTROL_BULTOS.NUMERO} = " + Me.txt_folio
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de entra de bultos"
            frmvistasprevias.Show 1
            
            MsgBox "Se a cerrado el movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
      End If
   Else
      If IsNumeric(Me.txt_folio) Then
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_control_bultos.rpt")
         reporte.RecordSelectionFormula = "{TB_ORACLE_CONTROL_BULTOS.NUMERO} = " + Me.txt_folio
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de entra de bultos"
         frmvistasprevias.Show 1
      Else
         MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_folio = ""
   Me.txt_fecha = Date
   If var_unidad_organizacional = 93 Then
      Me.txt_clave_destino = "CDI_ALMPT"
   End If
   Me.txt_nombre_destino = "PT. ALMACEN GENERAL"
   Me.txt_clave_origen = ""
   Me.txt_nombre_origen = ""
   Me.lv_bultos.ListItems.Clear
   Me.txt_total = ""
   Me.txt_clave_destino.Enabled = False
   Me.txt_nombre_destino.Enabled = False
   Me.txt_tipo_bulto.Enabled = True
   Me.txt_cantidad.Enabled = True
   Me.txt_clave_origen.Enabled = True
   Me.txt_clave_origen.SetFocus
   Me.txt_total = ""
   Me.txt_estatus = ""
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 1500
   Me.frm_busqueda.Visible = False
   Me.frm_lista_bultos.Visible = False
   Me.frm_lista_rutas.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_bultos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      var_Cantidad_posible = CDbl(Me.lv_bultos.selectedItem.SubItems(1))
      If var_Cantidad_posible > 0 Then
         If Me.txt_estatus = "" Then
            rs.Open "update tb_oracle_control_bultos set cantidad = cantidad - 1 where numero = " + Me.txt_folio + " and tipo_bulto = '" + Me.lv_bultos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_bultos.selectedItem.SubItems(1) = CDbl(Me.lv_bultos.selectedItem.SubItems(1)) - 1
            var_contador = 0
            For var_j = 1 To Me.lv_bultos.ListItems.Count
                Me.lv_bultos.ListItems.Item(var_j).Selected = True
                var_contador = var_contador + (Me.lv_bultos.selectedItem.SubItems(1))
            Next var_j
            Me.txt_total = var_contador
         Else
            MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Ya no puede ser modificado el movimiento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_lista_bultos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_bulto = Me.lv_lista_bultos.selectedItem
      Me.txt_tipo_bulto.SetFocus
   End If
End Sub

Private Sub lv_lista_bultos_LostFocus()
   Me.frm_lista_bultos.Visible = False
End Sub

Private Sub lv_lista_rutas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_clave_origen = Me.lv_lista_rutas.selectedItem
      Me.txt_nombre_origen = Me.lv_lista_rutas.selectedItem.SubItems(1)
      Me.txt_tipo_bulto.SetFocus
      Me.frm_lista_rutas.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.txt_clave_origen.SetFocus
   End If
End Sub

Private Sub lv_lista_rutas_LostFocus()
   'Me.frm_lista_rutas.Visible = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   
End Sub

Private Sub txt_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda) Then
         
         rs.Open "SELECT * FROM TB_ORACLE_CONTROL_BULTOS WHERE NUMERO = " + Me.txt_busqueda, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.lv_bultos.ListItems.Clear
            Me.txt_estatus = IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS)
            Me.txt_folio = Me.txt_busqueda
            Me.txt_fecha = Format(IIf(IsNull(rs!Fecha), "", rs!Fecha), "Short Date")
            Me.txt_clave_destino = rs!DESTINO
            Me.txt_clave_origen = rs!ORIGEN

            var_cadena = "SELECT * FROM XXVIA_VW_RUTAS WHERE CLAVE = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = var_cadena
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 40, Me.txt_clave_origen)
                 .Parameters.Append parametro
            End With
            Set rsaux7 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            Me.txt_nombre_origen = rsaux7!DESCRIPCION
            rsaux7.Close
            
            var_cadena = "SELECT * FROM XXVIA_VW_RUTAS WHERE CLAVE = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = var_cadena
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 40, Me.txt_clave_destino)
                 .Parameters.Append parametro
            End With
            Set rsaux7 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            Me.txt_nombre_destino = rsaux7!DESCRIPCION
            rsaux7.Close

            While Not rs.EOF
                  Set list_item = Me.lv_bultos.ListItems.Add(, , rs!TIPO_BULTO)
                  list_item.SubItems(1) = IIf(IsNull(rs!cantidad), "0", rs!cantidad)
                  rs.MoveNext
            Wend
            Me.txt_clave_origen.Enabled = False
            Me.txt_clave_destino.Enabled = False
            Me.txt_nombre_destino.Enabled = False
            Me.txt_nombre_origen.Enabled = False
            If Me.txt_estatus = "" Then
               Me.txt_tipo_bulto.Enabled = True
               Me.txt_tipo_bulto.SetFocus
            Else
               Me.txt_tipo_bulto.Enabled = False
            End If
            var_contador = 0
            For var_j = 1 To Me.lv_bultos.ListItems.Count
                Me.lv_bultos.ListItems.Item(var_j).Selected = True
                var_contador = var_contador + (Me.lv_bultos.selectedItem.SubItems(1))
            Next var_j
            Me.txt_total = var_contador
         
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.frm_busqueda.Visible = False
      Else
         MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_estatus = "" Then
         If IsNumeric(Me.txt_cantidad) Then
            If Me.txt_clave_destino <> "" Then
               If Me.txt_clave_origen <> "" Then
                  If Me.txt_folio = "" Then
                     rs.Open "SELECT ISNULL(MAX(NUMERO),0) AS FOLIO FROM TB_ORACLE_CONTROL_BULTOS ", cnn, adOpenDynamic, adLockOptimistic
                     Me.txt_folio = rs(0).Value + 1
                     rs.Close
                  End If
                  var_encontro = 0
                  For var_j = 1 To Me.lv_bultos.ListItems.Count
                      Me.lv_bultos.ListItems.Item(var_j).Selected = True
                      If Me.lv_bultos.selectedItem = Me.txt_tipo_bulto Then
                         var_encontro = var_j
                      End If
                  Next var_j
                  If var_encontro = 0 Then
                     rs.Open "INSERT INTO TB_ORACLE_CONTROL_BULTOS (NUMERO, FECHA, ESTATUS, ORIGEN, DESTINO, TIPO_BULTO, CANTIDAD, NOMBRE_ORIGEN, NOMBRE_DESTINO) VALUES (" + Me.txt_folio + ", GETDATE(), '', '" + Me.txt_clave_origen + "','" + Me.txt_clave_destino + "','" + Me.txt_tipo_bulto + "'," + Me.txt_cantidad + ",'" + Me.txt_nombre_origen + "','" + Me.txt_nombre_destino + "')", cnn, adOpenDynamic, adLockOptimistic
                     Set list_item = Me.lv_bultos.ListItems.Add(, , Me.txt_tipo_bulto)
                     list_item.SubItems(1) = Me.txt_cantidad
                  Else
                     rs.Open "UPDATE TB_ORACLE_CONTROL_BULTOS SET CANTIDAD = CANTIDAD +" + Me.txt_cantidad + " WHERE NUMERO = " + Me.txt_folio + " AND TIPO_BULTO = '" + Me.txt_tipo_bulto + "'", cnn, adOpenDynamic, adLockOptimistic
                     Me.lv_bultos.ListItems.Item(var_encontro).Selected = True
                     Me.lv_bultos.selectedItem.SubItems(1) = CDbl(Me.lv_bultos.selectedItem.SubItems(1)) + CDbl(Me.txt_cantidad)
                  End If
                  Me.txt_tipo_bulto = ""
                  Me.txt_cantidad = ""
                  Me.txt_tipo_bulto.SetFocus
                  var_contador = 0
                  For var_j = 1 To Me.lv_bultos.ListItems.Count
                      Me.lv_bultos.ListItems.Item(var_j).Selected = True
                      var_contador = var_contador + (Me.lv_bultos.selectedItem.SubItems(1))
                  Next var_j
                  Me.txt_total = var_contador
               Else
                  MsgBox "No se a seleccionado un origen", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a seleccionado un destino", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_clave_origen_GotFocus()
   Me.frm_lista_rutas.Visible = False
End Sub

Private Sub txt_clave_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista_rutas.ListItems.Clear
      var_cadena = "select * from xxvia_vw_rutas ORDER BY DESCRIPCION"
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = Me.lv_lista_rutas.ListItems.Add(, , rs!CLAVE)
            list_item.SubItems(1) = IIf(IsNull(rs!DESCRIPCION), "", rs!DESCRIPCION)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista_rutas.Visible = True
      'Me.lv_lista_rutas.SetFocus
      Me.txt_ruta_filtrar = ""
      Me.txt_ruta_filtrar.SetFocus
   End If
End Sub

Private Sub txt_clave_origen_LostFocus()
   If Me.txt_clave_origen <> "" Then
      Me.txt_clave_origen.Enabled = False
      Me.txt_nombre_origen.Enabled = False
   End If
End Sub

Private Sub txt_estatus_Change()
   If Me.txt_estatus = "I" Then
      Me.lbl_estatus.Caption = "Estatus: Cerrado"
   End If
   If Me.txt_estatus = "" Then
      Me.lbl_estatus.Caption = "Estatus: Abierto"
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ruta_filtrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_ruta_filtrar) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_ruta_filtrar)
             If Mid(Me.txt_ruta_filtrar, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_ruta_filtrar, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " where descripcion like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  descripcion like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  descripcion like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  descripcion like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  descripcion like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  descripcion like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  descripcion like '%" + var_like_7 + "%'"
      End If
      Me.lv_lista_rutas.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         rs.Open "SELECT * FROM xxvia_vw_rutas " + var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = Me.lv_lista_rutas.ListItems.Add(, , rs!CLAVE)
            list_item.SubItems(1) = IIf(IsNull(rs!DESCRIPCION), "", rs!DESCRIPCION)
            rs.MoveNext
         Wend
         rs.Close

      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_clave_origen.SetFocus
   End If
End Sub

Private Sub txt_tipo_bulto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista_bultos.ListItems.Clear
      rs.Open "select * from tb_oracle_empaques where empaque like '%COSTAL%' OR EMPAQUE = 'CAJA BIASI' ORDER BY EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = Me.lv_lista_bultos.ListItems.Add(, , rs!EMPAQUE)
            
            rs.MoveNext
      Wend
      Me.frm_lista_bultos.Visible = True
      Me.lv_lista_bultos.SetFocus
      rs.Close
  End If
End Sub

Private Sub txt_tipo_bulto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_tipo_bulto <> "" Then
         'Me.txt_cantidad = 1
         Me.txt_cantidad.SetFocus
      Else
         MsgBox "Tipo de bulto invalido", vbOKOnly, "ATENCION"
      End If
   Else
      KeyAscii = 0
   End If
End Sub
