VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbonificaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bonificaciones"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11640
   Begin VB.CommandButton cmd_nota_credito_electronica 
      Appearance      =   0  'Flat
      Caption         =   "NC Electronica"
      Enabled         =   0   'False
      Height          =   315
      Left            =   810
      Picture         =   "frmbonificaciones.frx":0000
      TabIndex        =   28
      Top             =   30
      Width           =   1485
   End
   Begin VB.Frame frm_lista2 
      Height          =   2400
      Left            =   2235
      TabIndex        =   25
      Top             =   780
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista2 
         Height          =   1875
         Left            =   30
         TabIndex        =   26
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3307
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
      Begin VB.Label lbl_lista2 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   1695
      Left            =   2850
      TabIndex        =   22
      Top             =   210
      Width           =   4050
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1230
         Left            =   30
         TabIndex        =   23
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   2170
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
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5380
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cargos "
      Height          =   5250
      Left            =   120
      TabIndex        =   12
      Top             =   1905
      Width           =   11355
      Begin VB.TextBox txt_total_aplicado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9735
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   4830
         Width           =   1545
      End
      Begin VB.Frame frm_cantidad_aplicar 
         Height          =   885
         Left            =   2460
         TabIndex        =   13
         Top             =   1320
         Width           =   2955
         Begin VB.TextBox txt_cantidad_aplicar 
            Height          =   360
            Left            =   1005
            TabIndex        =   14
            Top             =   390
            Width           =   1890
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            Caption         =   " Importe a Aplicar"
            ForeColor       =   &H8000000E&
            Height          =   270
            Left            =   0
            TabIndex        =   16
            Top             =   15
            Width           =   2940
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   165
            TabIndex        =   15
            Top             =   473
            Width           =   570
         End
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   4575
         Left            =   90
         TabIndex        =   18
         Top             =   225
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   8070
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
         MousePointer    =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Número"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Plazo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Moneda"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Abonos"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Aplicar    "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Clave Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Serie"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Iva"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmbonificaciones.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmbonificaciones.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11115
      Picture         =   "frmbonificaciones.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   11355
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   5190
      End
      Begin VB.TextBox txt_nombre_clase 
         Height          =   315
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   5190
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   1755
         TabIndex        =   4
         Top             =   255
         Width           =   1725
      End
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   10425
         TabIndex        =   9
         Top             =   255
         Width           =   795
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1755
         TabIndex        =   8
         Top             =   915
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1755
         TabIndex        =   6
         Top             =   585
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   9945
         TabIndex        =   20
         Top             =   315
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe a Aplicar:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   975
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   645
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   105
      TabIndex        =   19
      Top             =   315
      Width           =   11430
   End
End
Attribute VB_Name = "frmbonificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_numero_folio As Double
Dim var_numero_nota_inicio As Double
Dim var_serie As String
Dim var_clave_moneda As String
Dim var_tipo_Cambio As Double
Dim var_agente As String
Dim var_grupo_actual As String
Dim var_grupo_real As String
Dim var_titular As String
Dim var_plazo As Integer
Dim var_iva As Double
Dim var_numero_renglones As Double
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report



Private Sub nota_credito_turbina()
   Dim var_iva As String
   var_Archivo = App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat"
   Open (App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat") For Output As #2
   For var_k = var_numero_nota_inicio To var_numero_folio
       rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          'AQUI EMPIEZA LA NOTA DE CREDITO
          Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
          Print #1, Chr(15) + Chr(27) + Chr(64)
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          var_dia = Day(rs!dtim_Car_fecha)
          var_mes_numero = Month(rs!dtim_Car_fecha)
          If var_mes_numero = 1 Then
             var_mes = "ENERO"
          End If
          If var_mes_numero = 2 Then
             var_mes = "FEBRERO"
          End If
          If var_mes_numero = 3 Then
             var_mes = "MARZO"
          End If
          If var_mes_numero = 4 Then
             var_mes = "ABRIL"
          End If
          If var_mes_numero = 5 Then
             var_mes = "MAYO"
          End If
          If var_mes_numero = 6 Then
             var_mes = "JUNIO"
          End If
          If var_mes_numero = 7 Then
             var_mes = "JULIO"
          End If
          If var_mes_numero = 8 Then
             var_mes = "AGOSTO"
          End If
          If var_mes_numero = 9 Then
             var_mes = "SEPTIEMBRE"
          End If
          If var_mes_numero = 10 Then
             var_mes = "OCTUBRE"
          End If
          If var_mes_numero = 11 Then
             var_mes = "NOVIEMBRE"
          End If
          If var_mes_numero = 12 Then
             var_mes = "DICIEMBRE"
          End If
               
          var_año = Year(rs!dtim_Car_fecha)
          var_cadena = "                                                                                                           " + CStr(var_dia) + "   " + var_mes + "   " + CStr(var_año)
          Print #1, var_cadena
          Print #1, "    " + Str(rs!inte_Car_numero)
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          var_cliente = "CLIENTE: " + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
          var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
          var_cliente_sigo = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
          For var_j = 1 + Len(Trim(var_cliente)) To 83
              var_cliente = var_cliente + " "
          Next var_j
          var_cliente = var_cliente
          Print #1, Spc(5); var_cliente
          var_domicilio = "DOMICILIO: " + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " COLONIA: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
          var_agente = ""
          var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
          For var_j = 1 + Len(Trim(var_agente)) To 8
              var_agente = var_agente + " "
          Next var_j
          rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
          If Not rsaux4.EOF Then
             var_agente = var_agente + Mid(IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE), 1, 30)
          Else
             var_agente = var_agente + ""
          End If
          rsaux4.Close
          var_domicilio = var_domicilio
          Print #1, Spc(5); var_domicilio
          var_ciudad = ""
          var_ciudad = "CIUDAD: " + Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre))
          For var_j = 1 + Len(Trim(var_ciudad)) To 37
              var_ciudad = var_ciudad + " "
          Next var_j
          var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
          var_ciudad = var_ciudad
          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
          If Trim(var_rfc) <> "" Then
             var_ciudad = var_ciudad + " RFC: " + var_rfc
          Else
             var_ciudad = var_ciudad
          End If
            
          For var_j = 1 + Len((var_ciudad)) To 83
              var_ciudad = var_ciudad + " "
          Next var_j
                              
                               
          For var_j = 1 + Len(Trim(var_estado)) To 46
              var_estado = var_estado + " "
          Next var_j
                            
 
                                 
          var_ciudad = var_ciudad + "  " + var_agente
                              
          Print #1, Spc(5); var_ciudad
          var_rfc = "RFC:  " + var_rfc
          var_rfc = "ESTADO: " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
          For var_j = 1 + Len(Trim(var_rfc)) To 70
              var_rfc = var_rfc + " "
          Next var_j
          Print #1, Spc(5); var_rfc
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              
 ''' empieza el detalle
          rsaux.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_k), cnn, adOpenDynamic, adLockOptimistic
          var_contador_renglones_nota = 0
          While Not rsaux.EOF
                var_linea = ""
                var_linea = var_linea + " " + txt_clase + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux!inte_car_factura)
                If Len(Trim(var_linea)) < 115 Then
                   For var_j = 1 + Len(Trim(var_linea)) To 115
                       var_linea = var_linea + " "
                   Next var_j
                End If
                If Len(Trim(var_rfc)) = 0 Then
                   var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                Else
                   var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (1 + (rsaux!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                End If
                If Len(Trim(var_importe_str)) < 14 Then
                   For var_j = 1 + Len(Trim(var_importe_str)) To 14
                       var_importe_str = " " + var_importe_str
                   Next var_j
                End If
                var_linea = var_linea + var_importe_str
                Print #1, Spc(4); var_linea
                rsaux.MoveNext
                var_contador_renglones_nota = var_contador_renglones_nota + 1
          Wend
          rsaux.Close
          Print #1, ""
          Print #1, ""
          If var_contador_renglones_nota < var_numero_renglones Then
             For var_l = var_contador_renglones_nota To var_numero_renglones - 1
                 Print #1, ""
             Next var_l
          End If
                             
                             
                             
          If Len(Trim(var_rfc)) = 0 Then
             var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
             If Len(Trim(var_subimporte)) < 14 Then
                For var_j = 1 + Len(Trim(var_subimporte)) To 14
                    var_subimporte = " " + var_subimporte
                Next var_j
             End If
               
             var_linea = ""
             If Len(Trim(var_linea)) < 115 Then
                For var_j = 1 + Len(Trim(var_linea)) To 115
                    var_linea = var_linea + " "
                Next var_j
             End If
             var_linea = var_linea + var_subimporte
             Print #1, Spc(5); var_linea
                
             Print #1, ""
             var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
             If Len(Trim(var_linea)) < 115 Then
                For var_j = 1 + Len(Trim(var_linea)) To 115
                    var_linea = var_linea + " "
                Next var_j
             End If
                  
             var_iva = "-"
             For var_j = 1 + Len(Trim(var_iva)) To 14
                 var_iva = " " + var_iva
             Next var_j
                                    
             var_linea = var_linea + var_iva
             Print #1, Spc(5); var_linea
             Print #1, ""
                                   
             var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                
             If Len(Trim(var_importe)) < 14 Then
                For var_j = 1 + Len(Trim(var_importe)) To 14
                    var_importe = " " + var_importe
                Next var_j
             End If
             var_linea = ""
             If Len(Trim(var_linea)) < 115 Then
                For var_j = 1 + Len(Trim(var_linea)) To 115
                    var_linea = var_linea + " "
                Next var_j
             End If
             var_linea = var_linea + var_importe
             Print #1, Spc(5); var_linea
             
               
               
          Else
             var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
             
             If Len(Trim(var_subimporte)) < 14 Then
                For var_j = 1 + Len(Trim(var_subimporte)) To 14
                    var_subimporte = " " + var_subimporte
                Next var_j
             End If
             
             var_linea = ""
             If Len(Trim(var_linea)) < 115 Then
                For var_j = 1 + Len(Trim(var_linea)) To 115
                    var_linea = var_linea + " "
                Next var_j
             End If
             var_linea = var_linea + var_subimporte
             Print #1, Spc(5); var_linea
             
             Print #1, ""
             var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
             If Len(Trim(var_linea)) < 115 Then
                For var_j = 1 + Len(Trim(var_linea)) To 115
                    var_linea = var_linea + " "
                Next var_j
             End If
             
             var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
             
             For var_j = 1 + Len(Trim(var_iva)) To 14
                 var_iva = " " + var_iva
             Next var_j
             var_linea = var_linea + var_iva
             Print #1, Spc(5); var_linea
             Print #1, ""
             
             
             var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
             
             If Len(Trim(var_importe)) < 14 Then
                For var_j = 1 + Len(Trim(var_importe)) To 14
                    var_importe = " " + var_importe
                Next var_j
             End If
             var_linea = ""
             If Len(Trim(var_linea)) < 115 Then
                For var_j = 1 + Len(Trim(var_linea)) To 115
                    var_linea = var_linea + " "
                Next var_j
             End If
             var_linea = var_linea + var_importe
             Print #1, Spc(5); var_linea
               
          End If
          var_linea = ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Close #1
          Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
       End If
       rs.Close
   Next var_k
   Close #2
   x = Shell(var_Archivo, vbHide)
End Sub





Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmd_aplicar_Click()
End Sub

Private Sub cmd_imprimir_Click()
Dim var_importe_neto_1 As Double
Dim var_importe_total_1 As Double
Dim var_subimporte_1 As Double
Dim var_importe_iva_1 As Double

Dim var_tipo_Cambio As Double
Dim var_importe_factura As Double
Dim var_importe_pago As Double
Dim var_importe_saldo_pago As Double
Dim var_importe_total As Double
Dim var_fecha_pago As Date
Dim var_fecha_factura As Date
Dim var_contador_pagos As Double
Dim var_contador_facturas As Double
Dim var_descuento_agente As Double
Dim var_descuento_sistema As Double
Dim var_saldo As Double
Dim si As Integer
Dim i, n As Integer
Dim var_importe As Double
Dim var_descuento As Double
Dim var_importe_descuento As Double
Dim var_moneda_local As Integer
Dim var_posible_tipo_cambio As Boolean
Dim var_serie_cargo As String
Dim var_importe_neto As Double
Dim var_subimporte As Double
Dim var_importe_iva As Double
Dim var_k As Double
Dim var_l As Double
Dim var_numero_nota As Double
Dim var_contador_notas As Double
Dim var_saldo_real As Double
Dim var_iva_pasado As Double
var_posible_iva = 1
var_iva_pasado = 0
For var_j = 1 To lv_facturas.ListItems.Count
    lv_facturas.ListItems.item(var_j).Selected = True
    If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
       If var_iva_pasado = 0 Then
          var_iva_pasado = CDbl(Me.lv_facturas.selectedItem.SubItems(11))
       Else
          If var_iva_pasado <> CDbl(Me.lv_facturas.selectedItem.SubItems(11)) Then
             var_posible_iva = 0
          End If
       End If
    End If
Next var_j
If var_posible_iva = 1 Then
   var_iva = var_iva_pasado
If lv_facturas.ListItems.Count > 0 Then
   
   If Trim(txt_clase) <> "" Then
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
      End If
      rs.Close
      var_tipo_Cambio = 1
      If var_moneda_local = 0 Then
         rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
            var_posible_tipo_cambio = True
         Else
            var_posible_tipo_cambio = False
         End If
         rs.Close
      Else
         var_posible_tipo_cambio = True
      End If
      If var_posible_tipo_cambio = True Then
         var_contador_renglones = 0
         var_contador_notas = 0
         n = lv_facturas.ListItems.Count
         For i = 1 To n
            lv_facturas.ListItems.item(i).Selected = True
            If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
               var_contador_renglones = var_contador_renglones + 1
            End If
            If var_contador_renglones = var_numero_renglones Then
               var_contador_notas = var_contador_notas + 1
               var_contador_renglones = 0
            End If
         Next i
         If (var_contador_renglones > 0) And (var_contador_renglones < var_numero_renglones) Then
            var_contador_notas = var_contador_notas + 1
         End If
         rs.Open "select * from tb_Series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Ser_Serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_numero_folio = IIf(IsNull(rs!inte_ser_nota_credito), 1, rs!inte_ser_nota_credito)
         rs.Close
         var_numero_nota = var_numero_folio
         var_numero_nota_anterior = var_numero_nota
         var_numero_nota_inicio = var_numero_folio
         If var_contador_notas > 0 Then
            If var_contador_notas = 1 Then
               si = MsgBox("Se va a imprimir la nota de crédito número " + Str(var_numero_folio) + ", ¿la impresora esta lista?", vbYesNo, "ATENCION")
            End If
            If var_contador_notas > 1 Then
               si = MsgBox("Se van a imprimir de la nota " + Str(var_numero_folio) + " a la " + Str(var_numero_folio + (var_contador_notas - 1)) + ", ¿la impresora esta lista?", vbYesNo, "ATENCION")
            End If
            var_numero_nota_inicio = var_numero_folio
            If si = 6 Then
               si = MsgBox("Confirmar la impresión de la Nota de Crédito", vbYesNo, "ATENCION")
               If si = 6 Then
                  Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
                  Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
                  n = lv_facturas.ListItems.Count
                  
                  cnn.BeginTrans
                  var_contador_notas = 0
                  var_j = 0
                  
                  For i = 1 To n
                     lv_facturas.ListItems.item(i).Selected = True
                      If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                          var_j = var_j + 1
                      End If
                  Next i
                  For i = 1 To n
                     lv_facturas.ListItems.item(i).Selected = True
                     If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                        var_importe_neto_1 = ((lv_facturas.selectedItem.SubItems(8) * 1) * var_tipo_Cambio)
                        var_importe_total_1 = ((var_importe_neto_1 / (1 + (var_iva / 100))))
                        var_subimporte_1 = var_importe_total_1
                        var_importe_iva_1 = (var_importe_neto_1 - var_importe_total_1)
                         
                        var_importe_neto = var_importe_neto + var_importe_neto_1
                        var_importe_total = var_importe_total + var_importe_total_1
                        var_subimporte = var_subimporte + var_subimporte_1
                        var_importe_iva = var_importe_iva + var_importe_iva_1
                        
                        var_contador = var_contador + 1
                        
                        var_serie_cargo = lv_facturas.selectedItem.SubItems(10)
                        var_importe = lv_facturas.selectedItem.SubItems(8) * 1
                        var_descuento = 0
                        var_importe_descuento = 0
                        rs.Open "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_ecu_numero_cargo, vcha_ecu_serie_abono, vcha_ecu_movimiento_abono, inte_ecu_numero_abono, floa_ecu_importe_Cargo, floa_ecu_importe_abono) values ('" + var_empresa + "', '" + var_serie_cargo + "' ,'" + Trim(lv_facturas.selectedItem) + "', " + lv_facturas.selectedItem.SubItems(1) + ",'" + var_serie + "' ,'" + txt_clase + "'," + Str(var_numero_folio) + ", 0, " + Str(var_importe * var_tipo_Cambio) + ")", cnn, adOpenDynamic, adLockOptimistic
                        If var_empresa = "03" Then
                           rs.Open "SELECT * FROM VW_SALDOs_REALes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie_cargo + "' AND VCHA_CAR_DOCUMENTO = '" + Trim(lv_facturas.selectedItem) + "' AND INTE_CAR_NUMERO = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_saldo_real = IIf(IsNull(rs!SALDO), 0, rs!SALDO) - var_importe
                              rsaux6.Open "select * from tb_Saldos WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie_cargo + "' AND VCHA_CAR_DOCUMENTO = '" + Trim(lv_facturas.selectedItem) + "' AND INTE_CAR_NUMERO = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 If Round(rsaux6!FLOA_sAL_IMPORTE, 2) <> Round(rs!SALDO) Then
                                    If var_saldo_real < 0.01 Then
                                       var_saldo_real = 0
                                    End If
                                    rsaux5.Open "UPDATE TB_SALDOS SET FLOA_SAL_IMPORTE = " + CStr(var_saldo_real) + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie_cargo + "' AND VCHA_CAR_DOCUMENTO = '" + Trim(lv_facturas.selectedItem) + "' AND INTE_CAR_NUMERO = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                                 End If
                              End If
                              rsaux6.Close
                           End If
                           rs.Close
                        End If
                        rs.Open "insert into tb_detalle_bonificaciones (vcha_emp_empresa_id, vcha_car_documento,vcha_car_clase_id, vcha_ser_serie_id, inte_car_numero, inte_car_factura, floa_dbo_importe, floa_dbo_iva, char_dbo_estatus) values ('" + var_empresa + "', '" + txt_clase + "', '','" + var_serie + "', " + CStr(var_numero_folio) + "," + lv_facturas.selectedItem.SubItems(1) + ", " + Str(var_importe * var_tipo_Cambio) + "," + Str(var_iva) + ",'')"
                     End If
                     If (var_contador = var_numero_renglones) Or (i = n) Then
                        var_contador = 0
                        var_imprimir = True
                     End If
                     If var_imprimir = True Then
                        var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", txt_clase, txt_clase, var_numero_folio, "-", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", var_plazo, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                        rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_folio) + ", " + CStr(var_numero_folio) + ")", cnn, adOpenDynamic, adLockOptimistic
                        var_importe_neto = 0
                        var_importe_total = 0
                        var_subimporte = 0
                        var_importe_iva = 0
                        var_contador = 0
                        var_numero_folio = var_numero_folio + 1
                        var_contador_notas = var_contador_notas + 1
                     End If
                     var_imprimir = False
                  Next i
                  rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + " + CStr(var_contador_notas) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  
                  If var_empresa = "02" Or var_empresa = "03" Then
'''''''''''''''  IMPRESION DE LA NOTA DE CARGO
                     For var_k = var_numero_nota_inicio To var_numero_folio
                        rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                            'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                           Print #1, Chr(27) + Chr(64)
                           Print #1, Spc(92); Str(rs!inte_Car_numero)
                           Print #1, ""
                           Print #1, Spc(92); "       "; Format(rs!dtim_Car_fecha, "Short Date")
                           var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                           For var_j = 1 + Len(Trim(var_cliente)) To 83
                               var_cliente = var_cliente + " "
                           Next var_j
                           var_cliente = var_cliente + "AGUASCALIENTES, AGS."
                           Print #1, ""
                           Print #1, Spc(12); var_cliente
                           var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)) + " COL.: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + "  C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                           For var_j = 1 + Len(Trim(var_domicilio)) To 83
                               var_domicilio = var_domicilio + " "
                           Next var_j
                           var_agente = ""
                           var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                           For var_j = 1 + Len(Trim(var_agente)) To 8
                               var_agente = var_agente + " "
                           Next var_j
                           var_agente = var_agente + Mid(IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE), 1, 30)
                           var_domicilio = var_domicilio
                           Print #1, Spc(12); var_domicilio
                           var_ciudad = ""
                           var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                           For var_j = 1 + Len(Trim(var_ciudad)) To 37
                               var_ciudad = var_ciudad + " "
                           Next var_j
                           var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                           For var_j = 1 + Len(Trim(var_estado)) To 46
                               var_estado = var_estado + " "
                           Next var_j
                           var_ciudad = var_ciudad + var_estado
                                  
                           For var_j = 1 + Len(Trim(var_ciudad)) To 14
                               var_ciudad = var_ciudad + " "
                           Next var_j
                              
                           var_ciudad = var_ciudad + var_agente
                                
                           Print #1, Spc(12); var_ciudad
                           var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           var_rfc = "      " + var_rfc
                           For var_j = 1 + Len(Trim(var_rfc)) To 89
                               var_rfc = var_rfc + " "
                           Next var_j
                           var_rfc = var_rfc
                           Print #1, Spc(6); var_rfc
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           rsaux.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_k), cnn, adOpenDynamic, adLockOptimistic
                           var_contador_renglones_nota = 0
                           While Not rsaux.EOF
                              var_linea = txt_clase + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux!inte_car_factura)
                              If Len(Trim(var_linea)) < 120 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 120
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                              Else
                                 var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (1 + (rsaux!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                              End If
                              If Len(Trim(var_importe_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                     var_importe_str = " " + var_importe_str
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_str
                              Print #1, var_linea
                              rsaux.MoveNext
                              var_contador_renglones_nota = var_contador_renglones_nota + 1
                           Wend
                           rsaux.Close
                           If var_contador_renglones_nota < var_numero_renglones Then
                              For var_l = var_contador_renglones_nota To var_numero_renglones
                                  Print #1, ""
                              Next var_l
                           End If
                           var_cantidad_letra = rs!vcha_car_importe_letra
                           var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                           If Len(Trim(var_linea)) < 105 Then
                              For var_j = 1 + Len(Trim(var_linea)) To 105
                                  var_linea = var_linea + " "
                              Next var_j
                           End If
                           
                           
                           If Len(Trim(var_rfc)) = 0 Then
                              var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_subimporte_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                     var_subimporte_str = " " + var_subimporte_str
                                 Next var_j
                              End If
                              '1
                              var_iva_str = "-"
                              For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                  var_iva_str = " " + var_iva_str
                              Next var_j
                           Else
                              var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_subimporte_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                     var_subimporte_str = " " + var_subimporte_str
                                 Next var_j
                              End If
                              var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_iva_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                     var_iva_str = " " + var_iva_str
                                 Next var_j
                              End If
                           End If
                           var_linea = var_linea + "           " + var_subimporte_str
                           Print #1, Spc(4); var_linea
                           Print #1, Spc(120); var_iva_str
                           var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                           If Len(Trim(var_importe_str)) < 14 Then
                              For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                  var_importe_str = " " + var_importe_str
                              Next var_j
                           End If
                           Print #1, Spc(120); var_importe_str
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, Spc(85); "SISTEMAS"
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Print #1, ""
                           Close #1
                           
                           Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                           var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".bat"
                           Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                           Close #2
                           x = Shell(var_Archivo, vbHide)
                        End If
                        rs.Close
                     Next var_k
                  Else
'''''' impresion de notas de credito externas
                     If var_empresa = "16" Then
''''''''' multibondeados
                        var_Archivo = App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat"
                        Open (App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat") For Output As #2
                        For var_k = var_numero_nota_inicio To var_numero_folio
                           rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                              Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              Print #1, Chr(27) + Chr(64)
                              Print #1, Spc(92); Str(rs!inte_Car_numero)
                              Print #1, ""
                              Print #1, Spc(92); "       "; Format(rs!dtim_Car_fecha, "Short Date")
                              var_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente)) To 63
                                  var_cliente = var_cliente + " "
                              Next var_j
                              var_cliente = var_cliente + " "
                              Print #1, ""
                              Print #1, Spc(12); var_cliente
                              var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION))
                              var_j = 1 + Len(Trim(var_domicilio))
                              For var_j = var_j To 70
                                  var_domicilio = var_domicilio + " "
                              Next var_j
                              var_domicilio = var_domicilio + " AGUASCALIENTES, AGS"
                              var_j = Len(var_domicilio)
                              var_agente = ""
                              var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                              For var_j = 1 + Len(Trim(var_agente)) To 8
                                  var_agente = var_agente + " "
                              Next var_j
                              var_agente = var_agente
                              var_domicilio = var_domicilio
                              Print #1, Spc(12); var_domicilio
                              Print #1, Spc(12); IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                              var_ciudad = ""
                              var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              
                              For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                  var_ciudad = var_ciudad + " "
                              Next var_j
                                 
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_ciudad = var_ciudad
                              
                              For var_j = 1 + Len(Trim(var_rfc)) To 79
                                  var_rfc = var_rfc + " "
                              Next var_j
                              var_rfc = var_rfc + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              For var_j = 1 + Len(Trim(var_rfc)) To 103
                                  var_rfc = var_rfc + " "
                              Next var_j
                              var_rfc = var_rfc + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                              Print #1, Spc(12); var_ciudad
                              Print #1, Spc(12); var_rfc
                              Print #1, ""
                              Print #1, ""
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              rsaux.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_k), cnn, adOpenDynamic, adLockOptimistic
                              var_contador_renglones_nota = 0
                              While Not rsaux.EOF
                                 var_linea = ""
                                 var_linea = var_linea + " " + txt_clase + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux!inte_car_factura)
                                 If Len(Trim(var_linea)) < 105 Then
                                    For var_j = 1 + Len(Trim(var_linea)) To 105
                                        var_linea = var_linea + " "
                                    Next var_j
                                 End If
                                 If Len(Trim(var_rfc)) = 0 Then
                                    var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                 Else
                                    var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (1 + (rsaux!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                 End If
                                 If Len(Trim(var_importe_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                        var_importe_str = " " + var_importe_str
                                    Next var_j
                                 End If
                                 var_linea = var_linea + var_importe_str
                                 Print #1, Spc(4); var_linea
                                 rsaux.MoveNext
                                 var_contador_renglones_nota = var_contador_renglones_nota + 1
                              Wend
                              rsaux.Close
                              If var_contador_renglones_nota < var_numero_renglones Then
                                 For var_l = var_contador_renglones_nota To var_numero_renglones - 1
                                     Print #1, ""
                                 Next var_l
                              End If
                              
                              var_cantidad_letra = rs!vcha_car_importe_letra
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              If Len(Trim(var_linea)) < 91 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 91
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              
                              Print #1, ""
                              
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                        var_subimporte_str = " " + var_subimporte_str
                                    Next var_j
                                 End If
                                 
                                 var_iva_str = "-"
                                 For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                     var_iva_str = " " + var_iva_str
                                 Next var_j
                              Else
                                var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                       var_subimporte_str = " " + var_subimporte_str
                                   Next var_j
                                 End If
                                 var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_iva_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                        var_iva_str = " " + var_iva_str
                                    Next var_j
                                 End If
                              End If
                              var_linea = var_linea
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(8); var_linea
                              Print #1, ""
                              Print #1, Spc(110); var_subimporte_str
                              Print #1, ""
                              Print #1, Spc(110); var_iva_str
                              Print #1, ""
                              var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_importe_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                     var_importe_str = " " + var_importe_str
                                 Next var_j
                              End If
                              Print #1, Spc(110); var_importe_str
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Close #1
                              Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                           End If
                           rs.Close
                        Next var_k
                        Close #2
                        x = Shell(var_Archivo, vbHide)
                     
                     
                     
''''''''' fin de multibondeados
                        Else
                        
                        If var_empresa = "30" Then
                           Call nota_credito_turbina
                        Else
                           
                              var_Archivo = App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat"
                              Open (App.Path & "\nota_credito" + Trim(Str(var_numero_nota_inicio)) + ".bat") For Output As #2
                              For var_k = var_numero_nota_inicio To var_numero_folio
                                 rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                                 If Not rs.EOF Then
                               
                                  
                                    Open (App.Path & "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                                     'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                    Print #1, Chr(27) + Chr(64)
                                    Print #1, Spc(92); Str(rs!inte_Car_numero)
                                    Print #1, ""
                                    'Print #1, Spc(92); "       "; Format(rs!DTIM_CAR_FECHA, "Short Date")
                                    var_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    For var_j = 1 + Len(Trim(var_cliente)) To 63
                                        var_cliente = var_cliente + " "
                                    Next var_j
                                    var_cliente = var_cliente + " " + Format(rs!dtim_Car_fecha, "Short Date")
                                    Print #1, ""
                                    Print #1, Spc(12); var_cliente
                                    var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                    For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                        var_domicilio = var_domicilio + " "
                                    Next var_j
                                    var_agente = ""
                                    var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                    For var_j = 1 + Len(Trim(var_agente)) To 8
                                        var_agente = var_agente + " "
                                    Next var_j
                                    var_agente = var_agente
                                    var_domicilio = var_domicilio
                                    Print #1, Spc(12); var_domicilio
                                    var_ciudad = ""
                                    var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                                    var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                    If Len(Trim(var_estado)) > 0 Then
                                       var_ciudad = var_ciudad + ", " + var_estado
                                    End If
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                                    
                                    var_ciudad = var_ciudad
                                   
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    var_ciudad = var_ciudad + " " + var_rfc
                              
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 79
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                                    var_ciudad = var_ciudad + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 103
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                                    var_ciudad = var_ciudad + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                    
                                    Print #1, Spc(12); var_ciudad
                                 
                                    var_rfc = "      " + var_rfc
                                    For var_j = 1 + Len(Trim(var_rfc)) To 89
                                        var_rfc = var_rfc + " "
                                    Next var_j
                                    var_rfc = var_rfc
                                    'Print #1, Spc(6); var_rfc
                                    Print #1, ""
                                    Print #1, ""
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    rsaux.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_k), cnn, adOpenDynamic, adLockOptimistic
                                    var_contador_renglones_nota = 0
                                    While Not rsaux.EOF
                                       var_linea = ""
                                       var_linea = var_linea + " " + txt_clase + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux!inte_car_factura)
                                       If Len(Trim(var_linea)) < 105 Then
                                          For var_j = 1 + Len(Trim(var_linea)) To 105
                                              var_linea = var_linea + " "
                                          Next var_j
                                       End If
                                       If Len(Trim(var_rfc)) = 0 Then
                                          var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                       Else
                                          var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (1 + (rsaux!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                       End If
                                       If Len(Trim(var_importe_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                              var_importe_str = " " + var_importe_str
                                          Next var_j
                                       End If
                                       var_linea = var_linea + var_importe_str
                                       Print #1, var_linea
                                       rsaux.MoveNext
                                       var_contador_renglones_nota = var_contador_renglones_nota + 1
                                    Wend
                                    rsaux.Close
                                    If var_contador_renglones_nota < var_numero_renglones Then
                                       For var_l = var_contador_renglones_nota To var_numero_renglones - 1
                                           Print #1, ""
                                       Next var_l
                                    End If
                                    var_cantidad_letra = rs!vcha_car_importe_letra
                                    var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                    If Len(Trim(var_linea)) < 91 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 91
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                    
                                    Print #1, ""
                                    If Len(Trim(var_rfc)) = 0 Then
                                       var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                       If Len(Trim(var_subimporte_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                              var_subimporte_str = " " + var_subimporte_str
                                          Next var_j
                                       End If
                                       var_iva_str = "-"
                                       For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                           var_iva_str = " " + var_iva_str
                                       Next var_j
                                    Else
                                       var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                       If Len(Trim(var_subimporte_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                             var_subimporte_str = " " + var_subimporte_str
                                          Next var_j
                                       End If
                                       var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                       If Len(Trim(var_iva_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                              var_iva_str = " " + var_iva_str
                                          Next var_j
                                       End If
                                    End If
                                    var_linea = var_linea + "           " + var_subimporte_str
                                    Print #1, Spc(4); var_linea
                                    Print #1, Spc(106); var_iva_str
                                    var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                    If Len(Trim(var_importe_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                          var_importe_str = " " + var_importe_str
                                       Next var_j
                                    End If
                                    Print #1, Spc(106); var_importe_str
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, Spc(45); "SISTEMAS"
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, ""
                                    Print #1, ""
                                    Close #1
                                 
                                    Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                                 End If
                                 rs.Close
                              Next var_k
                              Close #2
                              x = Shell(var_Archivo, vbHide)
                           End If
                        End If
'''''' fin de impresion de notas de credito externas
                     End If
''''''''''
                     If Trim(txt_clave_cliente) <> "" Then
                        rs.Open "select * from tb_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
                           rs.Close
                           rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              lv_facturas.ListItems.Clear
                              var_contador_facturas = 0
                              While Not rs.EOF
                                 var_saldo = (IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
                                 If var_saldo > 0 Then
                                    Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_car_documento)
                                    var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                    var_importe_factura = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / var_tipo_Cambio
                                    list_item.SubItems(1) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                                    list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
                                    var_fecha_factura = Format(rs!dtim_Car_fecha, "Short Date")
                                    var_dias = var_fecha_pago - var_fecha_factura
                                    list_item.SubItems(3) = IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)
                                    list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                                    list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                                    list_item.SubItems(6) = Format(var_importe_factura - IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                                    list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                                    list_item.SubItems(8) = Format(0, "###,##0.00")
                                    list_item.SubItems(9) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                                    list_item.SubItems(10) = IIf(IsNull(rs!vcha_Ser_Serie_id), "", rs!vcha_Ser_Serie_id)
                                 End If
                                 rs.MoveNext:
                              Wend
                              rs.Close
                              txt_total_aplicado = Format(0, "###,##0.00")
                           Else
                              rs.Close
                              lv_facturas.ListItems.Clear
                              txt_total_aplicado = Format(0, "###,##0.00")
                           End If
                        Else
                           rs.Close
                           txt_clave_cliente = ""
                           txt_nombre_cliente = ""
                           txt_saldo = ""
                           txt_total_aplicado = Format(0, "###,##0.00")
                           lv_facturas.ListItems.Clear
                           MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
                        End If
                     End If
                     MsgBox "Se han terminado de aplicar los pagos", vbOKOnly, "ATENCION"
                   End If
                End If
              End If
         Else
            MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a indicado una clase del movimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El cliente seleccionado no tiene facturas vivas", vbOKOnly, "ATENCION"
   End If
Else
   MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
End If
   txt_clave_cliente.Enabled = False
   cmd_imprimir.Enabled = False
   txt_saldo.Enabled = False
End Sub

Private Sub cmd_nota_credito_electronica_Click()
   Dim var_importe_neto_1 As Double
   Dim var_importe_total_1 As Double
   Dim var_subimporte_1 As Double
   Dim var_importe_iva_1 As Double
   Dim var_tipo_Cambio As Double
   Dim var_importe_factura As Double
   Dim var_importe_pago As Double
   Dim var_importe_saldo_pago As Double
   Dim var_importe_total As Double
   Dim var_fecha_pago As Date
   Dim var_fecha_factura As Date
   Dim var_contador_pagos As Double
   Dim var_contador_facturas As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_saldo As Double
   Dim si As Integer
   Dim i, n As Integer
   Dim var_importe As Double
   Dim var_descuento As Double
   Dim var_importe_descuento As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_serie_cargo As String
   Dim var_importe_neto As Double
   Dim var_subimporte As Double
   Dim var_importe_iva As Double
   Dim var_k As Double
   Dim var_l As Double
   Dim var_numero_nota As Double
   Dim var_contador_notas As Double
   Dim var_saldo_real As Double
   Dim var_iva_pasado As Double
   var_posible_iva = 1
   var_iva_pasado = 0
   var_numero_renglones = 10000000
   For var_j = 1 To lv_facturas.ListItems.Count
       lv_facturas.ListItems.item(var_j).Selected = True
       If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
          If var_iva_pasado = 0 Then
             var_iva_pasado = CDbl(Me.lv_facturas.selectedItem.SubItems(11))
          Else
             If var_iva_pasado <> CDbl(Me.lv_facturas.selectedItem.SubItems(11)) Then
                var_posible_iva = 0
             End If
          End If
       End If
   Next var_j
   If var_posible_iva = 1 Then
      var_iva = var_iva_pasado
      If lv_facturas.ListItems.Count > 0 Then
         If Trim(txt_clase) <> "" Then
            If var_empresa = "02" Then
               If var_unidad_organizacional = "23" Then
                  var_serie = "NCEFT"
               Else
                  var_serie = "NCEMX"
               End If
            End If
            If var_empresa = "03" Then
               var_serie = "NCEVII"
            End If
            If var_empresa = "18" Then
               var_serie = "NCEVXX"
            End If
            If var_empresa = "30" Then
               var_serie = "NCETR"
            End If
                  
                  
            rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
            End If
            rs.Close
            var_tipo_Cambio = 1
            If var_moneda_local = 0 Then
               rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
                  var_posible_tipo_cambio = True
               Else
                  var_posible_tipo_cambio = False
               End If
               rs.Close
            Else
               var_posible_tipo_cambio = True
            End If
            If var_posible_tipo_cambio = True Then
               var_contador_renglones = 0
               var_contador_notas = 0
               n = lv_facturas.ListItems.Count
               For i = 1 To n
                   lv_facturas.ListItems.item(i).Selected = True
                   If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                      var_contador_renglones = var_contador_renglones + 1
                   End If
                   If var_contador_renglones = var_numero_renglones Then
                      var_contador_notas = var_contador_notas + 1
                      var_contador_renglones = 0
                   End If
               Next i
               If (var_contador_renglones > 0) And (var_contador_renglones < var_numero_renglones) Then
                  var_contador_notas = var_contador_notas + 1
               End If
               rs.Open "select * from tb_Series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Ser_Serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               var_numero_folio = IIf(IsNull(rs!inte_ser_nota_credito), 1, rs!inte_ser_nota_credito)
               rs.Close
               var_numero_nota = var_numero_folio
               var_numero_nota_anterior = var_numero_nota
               var_numero_nota_inicio = var_numero_folio
               If var_contador_notas > 0 Then
                  If var_contador_notas = 1 Then
                     si = MsgBox("Se va a imprimir la nota de crédito número " + Str(var_numero_folio) + ", ¿la impresora esta lista?", vbYesNo, "ATENCION")
                  End If
                  If var_contador_notas > 1 Then
                     si = MsgBox("Se van a imprimir de la nota " + Str(var_numero_folio) + " a la " + Str(var_numero_folio + (var_contador_notas - 1)) + ", ¿la impresora esta lista?", vbYesNo, "ATENCION")
                  End If
                  var_numero_nota_inicio = var_numero_folio
                  If si = 6 Then
                     si = MsgBox("Confirmar la impresión de la Nota de Crédito", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
                        Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
                        n = lv_facturas.ListItems.Count
                  
                        cnn.BeginTrans
                        var_contador_notas = 0
                        var_j = 0
                  
                        For i = 1 To n
                            lv_facturas.ListItems.item(i).Selected = True
                            If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                               var_j = var_j + 1
                            End If
                        Next i
                        For i = 1 To n
                            lv_facturas.ListItems.item(i).Selected = True
                            If (lv_facturas.selectedItem.SubItems(8) * 1) > 0 Then
                               var_importe_neto_1 = ((lv_facturas.selectedItem.SubItems(8) * 1) * var_tipo_Cambio)
                               var_importe_total_1 = ((var_importe_neto_1 / (1 + (var_iva / 100))))
                               var_subimporte_1 = var_importe_total_1
                               var_importe_iva_1 = (var_importe_neto_1 - var_importe_total_1)
                          
                               var_importe_neto = var_importe_neto + var_importe_neto_1
                               var_importe_total = var_importe_total + var_importe_total_1
                               var_subimporte = var_subimporte + var_subimporte_1
                               var_importe_iva = var_importe_iva + var_importe_iva_1
                        
                               var_contador = var_contador + 1
                        
                               var_serie_cargo = lv_facturas.selectedItem.SubItems(10)
                               var_importe = lv_facturas.selectedItem.SubItems(8) * 1
                               var_descuento = 0
                               var_importe_descuento = 0
                               rs.Open "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_ecu_numero_cargo, vcha_ecu_serie_abono, vcha_ecu_movimiento_abono, inte_ecu_numero_abono, floa_ecu_importe_Cargo, floa_ecu_importe_abono) values ('" + var_empresa + "', '" + var_serie_cargo + "' ,'" + Trim(lv_facturas.selectedItem) + "', " + lv_facturas.selectedItem.SubItems(1) + ",'" + var_serie + "' ,'" + txt_clase + "'," + Str(var_numero_folio) + ", 0, " + Str(var_importe * var_tipo_Cambio) + ")", cnn, adOpenDynamic, adLockOptimistic
                               If var_empresa = "03" Then
                                  rs.Open "SELECT * FROM VW_SALDOs_REALes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie_cargo + "' AND VCHA_CAR_DOCUMENTO = '" + Trim(lv_facturas.selectedItem) + "' AND INTE_CAR_NUMERO = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                                  If Not rs.EOF Then
                                     var_saldo_real = IIf(IsNull(rs!SALDO), 0, rs!SALDO) - var_importe
                                     rsaux6.Open "select * from tb_Saldos WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie_cargo + "' AND VCHA_CAR_DOCUMENTO = '" + Trim(lv_facturas.selectedItem) + "' AND INTE_CAR_NUMERO = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                                     If Not rsaux6.EOF Then
                                        If Round(rsaux6!FLOA_sAL_IMPORTE, 2) <> Round(rs!SALDO) Then
                                           'rsaux5.Open "UPDATE TB_SALDOS SET FLOA_SAL_IMPORTE = " + CStr(var_saldo_real) + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie_cargo + "' AND VCHA_CAR_DOCUMENTO = '" + Trim(lv_facturas.selectedItem) + "' AND INTE_CAR_NUMERO = " + lv_facturas.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                     End If
                                     rsaux6.Close
                                  End If
                                  rs.Close
                               End If
                               rs.Open "insert into tb_detalle_bonificaciones (vcha_emp_empresa_id, vcha_car_documento,vcha_car_clase_id, vcha_ser_serie_id, inte_car_numero, inte_car_factura, floa_dbo_importe, floa_dbo_iva, char_dbo_estatus) values ('" + var_empresa + "', '" + txt_clase + "', '','" + var_serie + "', " + CStr(var_numero_folio) + "," + lv_facturas.selectedItem.SubItems(1) + ", " + Str(var_importe * var_tipo_Cambio) + "," + Str(var_iva) + ",'')"
                            End If
                            If (var_contador = var_numero_renglones) Or (i = n) Then
                                var_contador = 0
                                var_imprimir = True
                            End If
                            If var_imprimir = True Then
                               var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", txt_clase, txt_clase, var_numero_folio, "-", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", var_plazo, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                               rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_folio) + ", " + CStr(var_numero_folio) + ")", cnn, adOpenDynamic, adLockOptimistic
                               var_importe_neto = 0
                               var_importe_total = 0
                               var_subimporte = 0
                               var_importe_iva = 0
                               var_contador = 0
                               var_numero_folio = var_numero_folio + 1
                               var_contador_notas = var_contador_notas + 1
                            End If
                            var_imprimir = False
                        Next i
                        rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + " + CStr(var_contador_notas) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
'''''''''''''''  IMPRESION DE LA NOTA DE CARGO
                        var_k = var_numero_nota_inicio
                        rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_k), cnn, adOpenDynamic, adLockOptimistic
                        If var_empresa = "28" Then
                           Dim dl As Long                                 ' Valor devuelto por la función API
                           Dim sAttributes As String                  ' Aributos
                           Dim sDriver As String                       ' Nombre del controlador
                           Dim sDescription As String                ' Descripción del DSN
                           Dim sDsnName As String                  ' Nombre del DSN
 
                           Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
                           Const vbAPINull As Long = 0&                         ' Puntero NULL

                            ' se elimina
                           Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
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
 
                           Set reporte = appl.OpenReport(App.Path + "\rep_nota_credito_vianney_catalog.rpt")
                           reporte.RecordSelectionFormula = "{VW_NOTA_CREDITO_VIANNEY_CATALOG.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_NOTA_CREDITO_VIANNEY_CATALOG.VCHA_SER_SERIE_ID} = '" + var_serie + "' and {VW_NOTA_CREDITO_VIANNEY_CATALOG.INTE_CAR_NUMERO} = " + CStr(var_k)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Nota de crédito"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                        Else
                           If Not rs.EOF Then
                              Open (App.Path & "\renombra" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                              Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
                              Close #2
                           
                              Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi") For Output As #1
                              'Open ("c:\NC_" + Trim(var_serie) + Trim(Str(rs!inte_car_numero)) + ".fi") For Output As #1
                              var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + rs!vcha_Ser_Serie_id + Chr(13) + "folio=" + CStr(rs!inte_Car_numero) + Chr(13)
                              var_año = CStr(Year(rs!dtim_Car_fecha))
                              var_mes = CStr(Month(rs!dtim_Car_fecha))
                              var_dia = CStr(Day(rs!dtim_Car_fecha))
                              var_hora = CStr(Hour(rs!dtim_Car_fecha))
                              var_minuto = CStr(Minute(rs!dtim_Car_fecha))
                              var_segundo = CStr(Second(rs!dtim_Car_fecha))
                              If Len(var_año) = 2 Then
                                 var_año = "20" + var_año
                              End If
                              If Len(var_mes) = 1 Then
                                 var_mes = "0" + var_mes
                              End If
                              If Len(var_dia) = 1 Then
                                 var_dia = "0" + var_dia
                              End If
                              If Len(var_hora) = 1 Then
                                 var_hora = "0" + var_hora
                              End If
                              If Len(var_minuto) = 1 Then
                                 var_minuto = "0" + var_minuto
                              End If
                              If Len(var_segundo) = 1 Then
                                 var_segundo = "0" + var_segundo
                              End If
                              
                              var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_rfc_cliente = ""
                              If var_rfc_cliente_1 = "" Then
                                 var_rfc_cliente = "XAXX010101000"
                              Else
                                 For var_j = 1 To Len(var_rfc_cliente_1)
                                     If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                        If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                           If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                              var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                           End If
                                        End If
                                     End If
                                 Next var_j
                              End If
                              
                              
                              var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                              var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
                              var_cadena = var_cadena + "noAprobacion=" + Chr(13)
                              var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
                              var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
                              var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
                              var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
                              If var_rfc_cliente = "XAXX010101000" Then
                                 var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                              Else
                                 var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_car_subimporte / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                              End If
                              var_cadena = var_cadena + "descuento=" + Chr(13)
                              var_cadena = var_cadena + "descuento1=" + Chr(13)
                              var_cadena = var_cadena + "descuento2=" + Chr(13)
                              var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
                              var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
                              var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
                              var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_certificado = rsaux2!vcha_emp_certificado
                              var_expedido = rsaux2!vcha_emp_expedido
                              If var_rfc_cliente = "XAXX010101000" Then
                                 var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,##0.000000") + Chr(13)
                              Else
                                 var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                              End If
                              var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                              var_cadena = var_cadena + "retencion=" + Chr(13)
                              var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
                              var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "<Emisor>" + Chr(13)
                              var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
                              var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
                              var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
                              var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
                              var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
                              var_cadena = var_cadena + "enoInterior=" + Chr(13)
                              var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
                              var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
                              var_cadena = var_cadena + "ereferencia=" + Chr(13)
                              var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
                              var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
                              var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
                              var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                              var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
                              var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
                              var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                              
                              var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "ex_calle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
                              var_cadena = var_cadena + "ex_noExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
                              var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
                              var_cadena = var_cadena + "ex_colonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
                              var_cadena = var_cadena + "ex_localidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
                              var_cadena = var_cadena + "ex_referencia=" + Chr(13)
                              var_cadena = var_cadena + "ex_municipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
                              var_cadena = var_cadena + "ex_estado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
                              var_cadena = var_cadena + "ex_pais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
                              var_cadena = var_cadena + "ex_codigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                              var_cadena = var_cadena + "</ExpedidoEn>"
                              
                              
                              
                              
                              var_cadena = var_cadena + "<Receptor>" + Chr(13)
                              var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
                              rsaux2.Close
                                            
                              var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_rfc_cliente = ""
                              If var_rfc_cliente_1 = "" Then
                                 var_rfc_cliente = "XAXX010101000"
                              Else
                                 For var_j = 1 To Len(var_rfc_cliente_1)
                                     If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                        If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                           If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                              var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                           End If
                                        End If
                                     End If
                                 Next var_j
                              End If
                              If var_empresa = "03" Or var_empresa = "28" Then
                                 var_rfc_cliente = "XEXX010101000"
                              End If
                              var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
                              var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
                              var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "<Cliente>" + Chr(13)
                              var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
                              var_cadena = var_cadena + "calle=" + Chr(13)
                              var_cadena = var_cadena + "noExterior=" + Chr(13)
                              var_cadena = var_cadena + "noInterior=" + Chr(13)
                              var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
                              var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + Chr(13)
                              rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
                              var_cadena = var_cadena + "referencia=" + Chr(13)
                              var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
                              var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
                              VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
                              If Trim(VAR_NOMBRE_PAIS) = "" Then
                                 VAR_NOMBRE_PAIS = "MEXICO"
                              End If
                              var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
                              var_cadena = var_cadena + Chr(13)
                              var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
                              var_cadena = var_cadena + "tel=" + Chr(13)
                              var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
                              var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                            
                              var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
                              var_cadena = var_cadena + "endomicilio=" + Chr(13)
                              var_cadena = var_cadena + "encalle=" + Chr(13)
                              var_cadena = var_cadena + "ennoExterior=" + Chr(13)
                              var_cadena = var_cadena + "ennoInterior=" + Chr(13)
                              var_cadena = var_cadena + "encolonia=" + Chr(13)
                              var_cadena = var_cadena + "enlocalidad=" + Chr(13)
                              var_cadena = var_cadena + "enreferencia=" + Chr(13)
                              var_cadena = var_cadena + "enmunicipio=" + Chr(13)
                              var_cadena = var_cadena + "enestado=" + Chr(13)
                              var_cadena = var_cadena + "enpais=" + Chr(13)
                              var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
                              var_cadena = var_cadena + "entel=" + Chr(13)
                              var_cadena = var_cadena + "enemail=" + Chr(13)
                              var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "<Concepto>" + Chr(13)
                              
                              
                              rsaux3.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_k), cnn, adOpenDynamic, adLockOptimistic
                              var_i = 1
                              While Not rsaux3.EOF
                                    pxx = CStr(var_i)
                                    If Len(pxx) = 1 Then
                                       pxx = "0" + pxx
                                    End If
                                    var_cadena = var_cadena + "p" + pxx + "_cantidad=1" + Chr(13)
                                    var_cadena = var_cadena + "p" + pxx + "_unidad=" + txt_clase + Chr(13)
                                    var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
                                    var_linea = txt_clase + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux3!inte_car_factura)
                                    var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                                    If var_rfc_cliente = "XAXX010101000" Then
                                       var_importe_str = ((IIf(IsNull(rsaux3!FLOA_dbo_IMPORTE), 0, rsaux3!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)))
                                    Else
                                       var_importe_str = ((IIf(IsNull(rsaux3!FLOA_dbo_IMPORTE), 0, rsaux3!FLOA_dbo_IMPORTE)) / (1 + (rsaux3!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)))
                                    End If
                                    var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                                    var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                                    rsaux3.MoveNext
                                    var_i = var_i + 1
                              Wend
                              rsaux3.Close
                              'MsgBox var_cadena
                              var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "<Otros>" + Chr(13)
                              var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
                              rs.MoveFirst
                              var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
                              var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
                              rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
                              rsaux1.Close
                              var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
                              var_cadena = var_cadena + "pedido=" + Chr(13)
                              var_cadena = var_cadena + "Embarque=" + Chr(13)
                              var_referencia_Bancaria = ""
                              var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
                              var_cadena = var_cadena + "fechaPedido=" + Chr(13)
                              var_cadena = var_cadena + "expedicion=" + Chr(13)
                              var_cadena = var_cadena + "observaciones=" + Chr(13)
                              var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
                              var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
                              var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
                              var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
                              var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                              
                              rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux11.EOF Then
                                 var_cadena = var_cadena + "agente=" + rsaux11!VCHA_AGE_AGENTE_ID + " " + rsaux11!VCHA_AGE_NOMBRE + Chr(13)
                              End If
                              rsaux11.Close
                              
                              If var_empresa = "02" Or var_empresa = "15" Or var_empresa = "16" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
                                 var_cadena = var_cadena + "formato=MHNCVTH_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "07" Then
                                 var_cadena = var_cadena + "formato=MHNCARE_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "31" Then
                                 var_cadena = var_cadena + "formato=MHNCCAN_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "42" Then
                                 var_cadena = var_cadena + "formato=MHNCCMA_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "41" Then
                                 var_cadena = var_cadena + "formato=MHNCCOP_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "150000" Then
                                 var_cadena = var_cadena + "formato=MHNCERE_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "33" Then
                                 var_cadena = var_cadena + "formato=MHNCMPU_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "34" Then
                                 var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "160000" Then
                                 var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "36" Then
                                 var_cadena = var_cadena + "formato=MHNCSME_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "30" Then
                                 var_cadena = var_cadena + "formato=MHNCTUR_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "44" Then
                                 var_cadena = var_cadena + "formato=MHNCUTV_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "38" Then
                                 var_cadena = var_cadena + "formato=MHNCVIA_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "40" Then
                                 var_cadena = var_cadena + "formato=MHNCVIN_V01.dat" + Chr(13)
                              End If
                              If var_empresa = "43" Then
                                 var_cadena = var_cadena + "formato=MHNCVOP_V01.dat" + Chr(13)
                              End If
                              
                              
                              var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "<addenda>" + Chr(13)
                              var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                              var_cadena = var_cadena + "</Factura>"
                              Print #1, var_cadena
                              Close #1
                              
                              var_Archivo = App.Path & "\renombra" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".bat"
                              x = Shell(var_Archivo, vbHide)
                           
'''''' fin de impresion de notas de credito externas
                           End If
                        End If
                        rs.Close
''''''''''
                        If Trim(txt_clave_cliente) <> "" Then
                           rs.Open "select * from tb_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
                              rs.Close
                              rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
                              If Not rs.EOF Then
                                 lv_facturas.ListItems.Clear
                                 var_contador_facturas = 0
                                 While Not rs.EOF
                                       var_saldo = (IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
                                       If var_saldo > 0 Then
                                          Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_car_documento)
                                          var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                          var_importe_factura = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / var_tipo_Cambio
                                          list_item.SubItems(1) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                                          list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
                                          var_fecha_factura = Format(rs!dtim_Car_fecha, "Short Date")
                                          var_dias = var_fecha_pago - var_fecha_factura
                                          list_item.SubItems(3) = IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)
                                          list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                                          list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                                          list_item.SubItems(6) = Format(var_importe_factura - IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                                          list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                                          list_item.SubItems(8) = Format(0, "###,##0.00")
                                          list_item.SubItems(9) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                                          list_item.SubItems(10) = IIf(IsNull(rs!vcha_Ser_Serie_id), "", rs!vcha_Ser_Serie_id)
                                       End If
                                       rs.MoveNext:
                                 Wend
                                 rs.Close
                                 txt_total_aplicado = Format(0, "###,##0.00")
                              Else
                                 rs.Close
                                 lv_facturas.ListItems.Clear
                                 txt_total_aplicado = Format(0, "###,##0.00")
                              End If
                           Else
                              rs.Close
                              txt_clave_cliente = ""
                              txt_nombre_cliente = ""
                              txt_saldo = ""
                              txt_total_aplicado = Format(0, "###,##0.00")
                              lv_facturas.ListItems.Clear
                              MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
                           End If
                        End If
                        MsgBox "Se han terminado de aplicar la nota de crédito", vbOKOnly, "ATENCION"
                      End If
                   End If
                 End If
            Else
               MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a indicado una clase del movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El cliente seleccionado no tiene facturas vivas", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No puede mezclar facturas con distintos tipos de IVA", vbOKOnly, "ATENCION"
   End If
   txt_clave_cliente.Enabled = False
   cmd_imprimir.Enabled = False
   txt_saldo.Enabled = False
End Sub

Private Sub cmd_nuevo_Click()
   txt_clave_cliente.Enabled = True
   If var_empresa = "15" Or var_empresa = "16" Or var_empresa = "31" Or var_empresa = "18" Or var_empresa = "03" Or var_empresa = "02" Or var_empresa = "06" Or var_empresa = "30" Or var_empresa = "28" Then
      cmd_imprimir.Enabled = False
      Me.cmd_nota_credito_electronica.Enabled = True
   Else
      cmd_imprimir.Enabled = True
      Me.cmd_nota_credito_electronica.Enabled = False
   End If
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   txt_saldo = ""
   lv_facturas.ListItems.Clear
   txt_total_aplicado = ""
   txt_saldo.Enabled = True
   If Me.txt_clase.Enabled = True Then
      txt_clase.SetFocus
   Else
      txt_clave_cliente.SetFocus
   End If
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   frm_lista2.Visible = False
   txt_clave_cliente.Enabled = False
   frm_lista.Visible = False
   txt_saldo.Enabled = False
   Top = 0
   Left = 0
   frm_cantidad_aplicar.Visible = False
   'rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   'var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   'rs.Close
   If var_empresa = "02" Or var_empresa = "03" Then
      var_numero_renglones = 38
   Else
      If var_empresa = "16" Then
         var_numero_renglones = 6
      Else
         If var_empresa = "30" Then
            var_numero_renglones = 10
         Else
            var_numero_renglones = 9
         End If
      End If
   End If
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      cmd_nuevo.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_Ser_Serie_id
      var_serie = rs!vcha_Ser_Serie_id
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      cmd_nuevo.Enabled = False
   End If
   rs.Close
   rs.Open "select * from tb_clases_Cartera where vcha_car_documento = 'BO' order by vcha_car_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
   If Not rs.EOF Then
      var_contador_movimiento = 0
      While Not rs.EOF
         var_contador_movimiento = var_contador_movimiento + 1
         rs.MoveNext
      Wend
      
      If var_contador_movimiento > 1 Then
         txt_nombre_clase.Enabled = True
         txt_clase.Enabled = True
      Else
         txt_nombre_clase.Enabled = False
         txt_clase.Enabled = False
      End If
      rs.MoveFirst
      txt_nombre_clase = rs!vcha_Car_nombre
      txt_clase = rs!vcha_Car_clase_id
   Else
      MsgBox "No se a indicado una clase de Bonificación", vbOKOnly, "ATENCION"
      txt_clase.Enabled = False
      cmb_clases.Enabled = False
   End If
   rs.Close
   If var_empresa = "15" Or var_empresa = "16" Or var_empresa = "31" Or var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "06" Or var_empresa = "30" Then
      Me.cmd_imprimir.Enabled = False
      Me.cmd_nota_credito_electronica.Enabled = True
   Else
      Me.cmd_imprimir.Enabled = True
      Me.cmd_nota_credito_electronica.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_bonificaciones)
End Sub

Private Sub lv_facturas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_facturas, ColumnHeader)
End Sub

Private Sub lv_facturas_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para indicar el importe a aplicar a la factura"
End Sub

Private Sub lv_facturas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If Trim(txt_saldo) = "" Then
         txt_saldo = 0
      End If
      If txt_saldo > 0 Then
         frm_cantidad_aplicar.Visible = True
         txt_cantidad_aplicar = ""
         txt_cantidad_aplicar.SetFocus
      Else
         MsgBox "No se a indicado el importe de la bonificación", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_facturas_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_clase = lv_lista.selectedItem
         txt_nombre_clase = lv_lista.selectedItem.SubItems(1)
         Me.txt_clase.SetFocus
      Else
         txt_clase = ""
         txt_nombre_clase = ""
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista2, ColumnHeader)
End Sub

Private Sub lv_lista2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista2.ListItems.Count > 0 Then
         txt_clave_cliente = lv_lista2.selectedItem
         txt_nombre_cliente = lv_lista2.selectedItem.SubItems(1)
      Else
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
      txt_clave_cliente.SetFocus
      frm_lista2.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista2.Visible = False
   End If
End Sub

Private Sub lv_lista2_LostFocus()
   frm_lista2.Visible = False
End Sub

Private Sub txt_cantidad_aplicar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_aplicar) Then
         If (txt_cantidad_aplicar * 1) + (lv_facturas.selectedItem.SubItems(8) * 1) > (lv_facturas.selectedItem.SubItems(7) * 1) Then
            MsgBox "La cantidad a aplicar exede el importe del saldo de la factura", vbOKOnly, "ATENCIO"
         Else
            If (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1) <= (txt_saldo * 1) Then
               lv_facturas.selectedItem.SubItems(8) = Format(txt_cantidad_aplicar + (lv_facturas.selectedItem.SubItems(8) * 1), "###,##0.00")
               txt_total_aplicado = (txt_total_aplicado * 1) + (txt_cantidad_aplicar * 1)
            Else
               MsgBox "La cantidad a aplicar exede al importe del saldo del pago del cliente", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      End If
      frm_cantidad_aplicar.Visible = False
      lv_facturas.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_cantidad_aplicar.Visible = False
      If lv_facturas.ListItems.Count > 0 Then
         lv_facturas.SetFocus
      End If
   End If
End Sub

Private Sub txt_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BO' or vcha_Car_documento = 'BF' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Me.txt_clase <> "" Then
      'MsgBox "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BO' or vcha_Car_documento = 'BF and vcha_Car_clase = '" + Me.txt_clase + "'"
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where (vcha_car_documento= 'BO' or vcha_Car_documento = 'BF') and vcha_Car_clase_ID = '" + Me.txt_clase + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_clase = IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre)
      Else
         MsgBox "Clase de cartera no existe", vbOKOnly, "ATENCION"
         Me.txt_clase = ""
         Me.txt_nombre_clase = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_clase = ""
   End If
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select * from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      var_tipo_lista = 1
      frm_lista2.Visible = True
      lv_lista2.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Dim var_importe_factura As Double
   Dim var_importe_pago As Double
   Dim var_importe_saldo_pago As Double
   Dim var_importe_total As Double
   Dim var_fecha_pago As Date
   Dim var_fecha_factura As Date
   Dim var_contador_pagos As Double
   Dim var_contador_facturas As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_saldo As Double
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id ='" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
         var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         'var_iva = 15
         rs.Close
         rs.Open "select * from vw_saldos_facturas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente + "' and floa_sal_importe > 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lv_facturas.ListItems.Clear
            var_contador_facturas = 0
            While Not rs.EOF
               var_saldo = (IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE) - IIf(IsNull(rs!importe_saldo), 0, rs!importe_saldo))
               If var_saldo > 0 Then
                  Set list_item = lv_facturas.ListItems.Add(, , rs!vcha_car_documento)
                  var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                  var_importe_factura = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / var_tipo_Cambio
                  list_item.SubItems(1) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                  list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
                  var_fecha_factura = Format(rs!dtim_Car_fecha, "Short Date")
                  var_dias = var_fecha_pago - var_fecha_factura
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_CAR_PLAZO), 0, rs!INTE_CAR_PLAZO)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                  list_item.SubItems(5) = Format(var_importe_factura, "###,##0.00")
                  list_item.SubItems(6) = Format(var_importe_factura - IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,##0.00")
                  list_item.SubItems(8) = Format(0, "###,##0.00")
                  list_item.SubItems(9) = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                  list_item.SubItems(10) = IIf(IsNull(rs!vcha_Ser_Serie_id), "", rs!vcha_Ser_Serie_id)
                  'list_item.SubItems(11) = IIf(IsNull(rs!floa_car_porcentaje_iva), "", rs!floa_car_porcentaje_iva)
                  list_item.SubItems(11) = var_iva
               End If
               rs.MoveNext:
            Wend
            rs.Close
            txt_total_aplicado = Format(0, "###,##0.00")
         Else
            rs.Close
            lv_facturas.ListItems.Clear
            txt_total_aplicado = Format(0, "###,##0.00")
         End If
      Else
         rs.Close
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         txt_saldo = ""
         txt_total_aplicado = Format(0, "###,##0.00")
         lv_facturas.ListItems.Clear
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If Me.lv_facturas.ListItems.Count > 20 Then
      lv_facturas.ColumnHeaders(2).Width = 1250
   Else
      lv_facturas.ColumnHeaders(2).Width = 1500.09
   End If
End Sub

Private Sub txt_descuento_Change()

End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_nombre_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'BO' or vcha_car_documento = 'BF' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista2.ListItems.Clear
      rs.Open "select * from vw_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista2.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista2 = "CLIENTES"
      var_tipo_lista = 1
      frm_lista2.Visible = True
      lv_lista2.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_saldo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.lv_facturas.ListItems.Count > 0 Then
         lv_facturas.SetFocus
      End If
   End If
End Sub
