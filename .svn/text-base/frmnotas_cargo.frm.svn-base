VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnotas_cargo 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas de Cargo"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8130
   Begin VB.Frame frm_lista 
      Height          =   2565
      Left            =   1125
      TabIndex        =   19
      Top             =   60
      Width           =   6390
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2100
         Left            =   45
         TabIndex        =   20
         Top             =   405
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   3704
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   6315
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmnotas_cargo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmnotas_cargo.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7665
      Picture         =   "frmnotas_cargo.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   60
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   2040
      Left            =   120
      TabIndex        =   8
      Top             =   465
      Width           =   7920
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   885
         TabIndex        =   6
         Top             =   1605
         Width           =   795
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1695
         TabIndex        =   7
         Top             =   1605
         Width           =   1530
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   570
         Width           =   5175
      End
      Begin VB.TextBox txt_nombre_clase_cartera 
         Height          =   315
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   255
         Width           =   5175
      End
      Begin VB.TextBox txt_clase_cartera 
         Height          =   315
         Left            =   885
         TabIndex        =   0
         Top             =   255
         Width           =   1725
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   1275
         Width           =   1725
      End
      Begin VB.TextBox txt_plazo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   4
         Top             =   945
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   885
         TabIndex        =   2
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lbl_moneda 
         Height          =   285
         Left            =   2670
         TabIndex        =   16
         Top             =   1305
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   660
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   105
      TabIndex        =   15
      Top             =   330
      Width           =   7920
   End
End
Attribute VB_Name = "frmnotas_cargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim var_plazo As Integer
   Dim si, i, n As Integer
   Dim var_serie As String
   Dim VAR_TIPO_LISTA As Integer



Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmb_series_KeyPress(KeyAscii As Integer)
   Me.cmd_imprimir.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim si As Integer
   Dim var_importe_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_neto As Double
   Dim var_subimporte As Double
   Dim var_tipo_Cambio As Double
   Dim var_numero_folio As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_j As Integer
   Dim var_subimporte_str As String
   Dim var_importe_str As String
   Dim var_iva_str As String
   Dim var_ciudad As String
   Dim var_rfc As String
   Dim var_linea As String
   Dim var_estado As String
   
   var_moneda_local = 1
   If Trim(txt_serie) <> "" Then
   If Trim(txt_numero) <> "" Then
   If IsNumeric(txt_numero) Then
   If Trim(Me.txt_clave_cliente) <> "" Then
      If Trim(Me.txt_clase_cartera) <> "" Then
        If IsNumeric(Me.txt_importe) Then
            var_moneda_local = 1
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
               si = MsgBox("¿Deseas imprimir la Nota de Cargo " + txt_numero + "?", vbYesNo, "ATENCION")
               If si = 6 Then
                  si = MsgBox("Confirmar la impresíon de la Nota de Cargo " + txt_numero, vbYesNo, "ATENCION")
                  If si = 6 Then
                     var_serie = txt_serie
                     var_numero_folio = CDbl(txt_numero)
                     rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'NG' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND INTE_CAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                     If rs.EOF Then
                     rs.Close
                     MsgBox "Se va a imprimir la Nota de cargo Número " + Str(var_numero_folio), vbYesNo, "ATENCION"
                     si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
                     If si = 6 Then
                        cnn.BeginTrans
                        var_importe_total = (txt_importe / (1 + (var_iva / 100))) * var_tipo_Cambio
                        var_importe_neto = txt_importe * var_tipo_Cambio
                        var_importe_iva = var_importe_neto - var_importe_total
                        var_subimporte = var_importe_total
                        var_insertar = False
                        var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NG", "NC", txt_clase_cartera, var_numero_folio, "+", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", var_plazo, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                        var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "NC", var_numero_folio, "", "", 0, var_importe_neto, 0)
                        rsaux3.Open "update tb_principal set inte_pri_nota_cargo = inte_pri_nota_cargo + 1", cnn, adOpenDynamic, adLockOptimistic
                        rsaux3.Open "update tb_series set inte_ser_nota_cargo =  inte_ser_nota_cargo + 1 where vcha_emp_empresa_id= '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
                        
                        rs.Open "select * from vw_notas_cargo where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'NC' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
            ''''''''' '''  IMPRESION DE LA NOTA DE CARGO
                           If var_empresa = "16" Then
''''' nota de cargo para otras empresas
                              var_Archivo = App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt"
                              Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                              Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              Print #1, Chr(27) + Chr(64)
                              Print #1, Spc(92); Str(rs!inte_Car_numero)
                              Print #1, ""
                              Print #1, Spc(92); "       "; Format(rs!DTIM_CAR_FECHA, "Short Date")
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
                              var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
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
                              var_rfc = var_rfc + IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
                              Print #1, Spc(12); var_ciudad
                              Print #1, Spc(12); var_rfc
                              Print #1, ""
                              Print #1, ""
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           
                              var_linea = "NC" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre
                        
                              If Len(Trim(var_linea)) < 108 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 108
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_importe_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              Else
                                 var_importe_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) / (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                              End If
                              If Len(Trim(var_importe_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                     var_importe_str = " " + var_importe_str
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_str
                              Print #1, var_linea
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                        
                        
                        
                              var_cantidad_letra = rs!vcha_car_importe_letra
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              If Len(Trim(var_linea)) < 91 Then
                                For var_j = 1 + Len(Trim(var_linea)) To 91
                                    var_linea = var_linea + " "
                                 Next var_j
                              End If
                       
                              Print #1, ""
                         
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                        var_subimporte_str = " " + var_subimporte_str
                                     Next var_j
                                 End If
                                 var_iva_str = "      -        "
                                 For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                    var_iva_str = " " + var_iva_str
                                 Next var_j
                              Else
                                 var_subimporte_str = Format(((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
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
                              var_importe_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
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
                              Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                              var_Archivo = App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".bat"
                              Print #2, "copy " + App.Path + "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                              Close #2
                              x = Shell(var_Archivo, vbHide)
                           
                           Else
                              Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                              Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              Print #1, Spc(92); Str(rs!inte_Car_numero)
                              Print #1, ""
                              Print #1, ""
                             'Print #1, Spc(70); Format(rs!DTIM_CAR_FECHA, "Short Date")
                              var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente)) To 73
                                  var_cliente = var_cliente + " "
                              Next var_j
                              var_cliente = var_cliente + Format(rs!DTIM_CAR_FECHA, "Short Date")
                              Print #1, Spc(12); var_cliente
                              var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_domicilio)) To 73
                                  var_domicilio = var_domicilio + " "
                              Next var_j
                              var_agente = ""
                              'var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                              'For var_j = 1 + Len(Trim(var_agente)) To 8
                              '    var_agente = var_agente + " "
                              'Next var_j
                             ' var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                              var_domicilio = var_domicilio
                              Print #1, Spc(5); var_domicilio
                              var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                  var_ciudad = var_ciudad + " "
                              Next var_j
                              'var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              'For var_j = 1 + Len(Trim(var_estado)) To 46
                              '    var_estado = var_estado + " "
                              'Next var_j
                              'var_ciudad = var_ciudad + var_estado
                           
                              'For var_j = 1 + Len(Trim(var_ciudad)) To 14
                              '    var_ciudad = var_ciudad + " "
                              'Next var_j
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_rfc = "RFC:  " + var_rfc
                              'For var_j = 1 + Len(Trim(var_rfc)) To 89
                              '    var_rfc = var_rfc + " "
                              'Next var_j
                              var_rfc = var_rfc
                        
                              var_ciudad = var_ciudad + var_rfc
                         
                              Print #1, Spc(5); var_ciudad
                              Print #1, ""
                              Print #1, ""
                              var_linea = "NC" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre
                              If Len(Trim(var_linea)) < 108 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 108
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_importe_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              Else
                                 var_importe_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) / (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                              End If
                              If Len(Trim(var_importe_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                     var_importe_str = " " + var_importe_str
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_str
                              Print #1, var_linea
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, ""
                              var_cantidad_letra = rs!vcha_car_importe_letra
                        
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              If Len(Trim(var_linea)) < 93 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 93
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                        var_subimporte_str = " " + var_subimporte_str
                                    Next var_j
                                 End If
                                 var_iva_str = "      -        "
                                 For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                     var_iva_str = " " + var_iva_str
                                 Next var_j
                               Else
                                 var_subimporte_str = Format(((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
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
                              Print #1, Spc(108); var_iva_str
                              var_importe_str = Format((IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              If Len(Trim(var_importe_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                     var_importe_str = " " + var_importe_str
                                 Next var_j
                              End If
                              Print #1, Spc(108); var_importe_str
                              Print #1, ""
                              'Print #1, Spc(4); "ESTA DOCUMENTO SERA PAGADO EN UNA SOLA EXHIBICION"
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(85); "SISTEMAS"
                              Close #1
                    
                              Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                              var_Archivo = App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".bat"
                              Print #2, "copy " + App.Path + "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                              Close #2
                              x = Shell(var_Archivo, vbHide)
                           End If
'''''''''
                           txt_clase_cartera.Enabled = True
                           txt_nombre_clase_cartera.Enabled = True
                           txt_plazo.Enabled = True
                           txt_importe.Enabled = True
                           cmb_clientes = ""
                           txt_clave_cliente = ""
                           txt_plazo = ""
                           txt_importe = ""
                           txt_clave_empresa = ""
                           lbl_moneda = ""
                           Me.txt_nombre_clase_cartera = ""
                           Me.txt_clase_cartera = ""
                           Me.txt_nombre_cliente = ""

                        End If
                        If rs.State = 1 Then
                           rs.Close
                        End If
                     Else
                       MsgBox "La impresión de la Nota de Cargo a sido cancelada", vbOKOnly, "ATENCION"
                     End If
                     Else
                        rs.Close
                        MsgBox "La nota de cargo ya existe", vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            Else
               MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de cartera incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
   Else
      MsgBox "Numero de documento incorecto", vbOKOnly, "ATENCION"
   End If
   Else
      MsgBox "Se debe de indicar un numero de documento", vbOKOnly, "ATENCION"
   End If
   Else
      MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_clase_cartera.Enabled = True
   txt_nombre_clase_cartera.Enabled = True
   txt_plazo.Enabled = True
   txt_importe.Enabled = True
   cmb_clientes = ""
   txt_clave_cliente = ""
   txt_plazo = ""
   txt_importe = ""
   txt_clave_empresa = ""
   lbl_moneda = ""
   Me.txt_nombre_clase_cartera = ""
   Me.txt_clase_cartera = ""
   Me.txt_nombre_cliente = ""
   txt_clase_cartera.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 2500
   Left = 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_notas_cargo)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If VAR_TIPO_LISTA = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clase_cartera = lv_lista.selectedItem
            txt_nombre_clase_cartera = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clase_cartera = ""
            txt_nombre_clase_cartera = ""
         End If
         frm_lista.Visible = False
         txt_clase_cartera.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_cliente = ""
            txt_nombre_cliente = ""
         End If
         frm_lista.Visible = False
         txt_clave_cliente.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
      If VAR_TIPO_LISTA = 1 Then
         txt_clase_cartera.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         txt_clave_cliente.SetFocus
      End If
   End If
End Sub



Private Sub lv_lista_LostFocus()
    frm_lista.Visible = False
End Sub

Private Sub txt_clase_cartera_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clase_cartera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      VAR_TIPO_LISTA = 1
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'NC' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4799.74
      Else
         lv_lista.ColumnHeaders(2).Width = 4999.74
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_cartera_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clase_cartera_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clase_cartera) <> "" Then
      rs.Open "select * from tb_clases_cartera where vcha_car_clase_id = '" + txt_clase_cartera + "' and vcha_car_documento = 'NC'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_clase_cartera = rs!vcha_Car_nombre
      Else
         MsgBox "Clave de clase de cartera incorrecta", vbOKOnly, "ATENCION"
         txt_clase_cartera = ""
         txt_nombre_clase_cartera = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_clave_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      VAR_TIPO_LISTA = 2
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_Actual_id is not null order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4799.74
      Else
         lv_lista.ColumnHeaders(2).Width = 4999.74
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_Actual_id is not null", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_cliente = txt_clave_cliente
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         txt_plazo = 0
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         lbl_moneda = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
      Else
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         var_grupo_actual = ""
         var_grupo_real = ""
         var_cliente = ""
         var_titular = ""
         var_clave_moneda = ""
         var_agente = ""
         var_plazo = 0
         txt_plazo = var_plazo
         var_iva = 0
         var_clave_moneda = ""
         lbl_moneda = ""
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_clave_cliente = ""
      txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_folio_LostFocus()
   If Not IsNumeric(txt_folio) Then
      MsgBox "Folio Incorrecto", vbOKOnly, "ATENCION"
      txt_folio = ""
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_LostFocus()
   If Not IsNumeric(txt_importe) Then
      MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      txt_importe = 0
   End If
End Sub

Private Sub txt_nombre_clase_cartera_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_clase_cartera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      VAR_TIPO_LISTA = 1
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'NC' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre))
            rs.MoveNext
         Wend
      End If
      rs.Close
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4799.74
      Else
         lv_lista.ColumnHeaders(2).Width = 4999.74
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_cartera_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_clase_cartera_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      VAR_TIPO_LISTA = 2
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_Actual_id is not null  order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4799.74
      Else
         lv_lista.ColumnHeaders(2).Width = 4999.74
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_plazo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_plazo_LostFocus()
   If IsNumeric(txt_plazo) Then
      var_plazo = txt_plazo
   Else
      MsgBox "Plazo incorrecto", vbOKOnly, "ATENCION"
      txt_plazo = var_plazo
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
