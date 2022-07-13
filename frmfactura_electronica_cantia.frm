VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfactura_electronica_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación electrónica CANTIA"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmfactura_electronica_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_factura_electronica 
      Caption         =   "Factura electrónica"
      Height          =   345
      Left            =   90
      TabIndex        =   32
      Top             =   0
      Width           =   1605
   End
   Begin VB.Frame Frame4 
      Caption         =   " Datos del cliente "
      Height          =   2475
      Left            =   75
      TabIndex        =   10
      Top             =   4770
      Width           =   11490
      Begin VB.TextBox txt_correo 
         Height          =   390
         Left            =   6750
         TabIndex        =   30
         Top             =   1935
         Width           =   4635
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   390
         Left            =   870
         TabIndex        =   28
         Top             =   1935
         Width           =   4635
      End
      Begin VB.TextBox txt_municipio 
         Enabled         =   0   'False
         Height          =   390
         Left            =   6750
         TabIndex        =   26
         Top             =   1515
         Width           =   4635
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   390
         Left            =   870
         TabIndex        =   24
         Top             =   1515
         Width           =   4635
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   390
         Left            =   6750
         TabIndex        =   22
         Top             =   1095
         Width           =   4635
      End
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   390
         Left            =   870
         TabIndex        =   20
         Top             =   1095
         Width           =   4635
      End
      Begin VB.TextBox txt_direccion 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4020
         TabIndex        =   18
         Top             =   675
         Width           =   7380
      End
      Begin VB.TextBox txt_cp 
         Enabled         =   0   'False
         Height          =   390
         Left            =   870
         TabIndex        =   16
         Top             =   675
         Width           =   2055
      End
      Begin VB.TextBox txt_nombre 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4020
         TabIndex        =   14
         Top             =   255
         Width           =   7380
      End
      Begin VB.TextBox txt_rfc 
         Enabled         =   0   'False
         Height          =   390
         Left            =   870
         TabIndex        =   12
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Correo:"
         Height          =   195
         Left            =   5970
         TabIndex        =   29
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Left            =   5970
         TabIndex        =   25
         Top             =   1613
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   1613
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   5970
         TabIndex        =   21
         Top             =   1193
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1193
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   3090
         TabIndex        =   17
         Top             =   773
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   773
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   3060
         TabIndex        =   13
         Top             =   353
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   353
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Detalle "
      Height          =   3150
      Left            =   75
      TabIndex        =   9
      Top             =   1590
      Width           =   11460
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
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   7950
         TabIndex        =   6
         Top             =   2595
         Width           =   3375
      End
      Begin MSComctlLib.ListView lv_detalle 
         Height          =   2295
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   4048
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "%D"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Descuento"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Precio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   1852
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL:"
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
         Left            =   6720
         TabIndex        =   31
         Top             =   2670
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   45
      TabIndex        =   8
      Top             =   375
      Width           =   11565
   End
   Begin VB.Frame Frame1 
      Caption         =   " Folio "
      Height          =   1080
      Left            =   75
      TabIndex        =   7
      Top             =   465
      Width           =   11430
      Begin VB.CheckBox chk_factura_del_dia 
         Caption         =   "Factura del dia"
         Height          =   315
         Left            =   9585
         TabIndex        =   34
         Top             =   480
         Width           =   1395
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "Buscar factura"
         Height          =   510
         Left            =   7380
         TabIndex        =   4
         Top             =   390
         Width           =   1890
      End
      Begin VB.TextBox txt_tipo_4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   4920
         TabIndex        =   3
         Top             =   330
         Width           =   2400
      End
      Begin VB.TextBox txt_tipo_3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   4110
         TabIndex        =   2
         Top             =   330
         Width           =   765
      End
      Begin VB.TextBox txt_tipo_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3375
         TabIndex        =   1
         Top             =   330
         Width           =   675
      End
      Begin VB.TextBox txt_tipo_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2565
         TabIndex        =   0
         Top             =   330
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmfactura_electronica_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_buscar_Click()
   If IsNumeric(Me.txt_tipo_1) Then
      If IsNumeric(Me.txt_tipo_2) Then
         If IsNumeric(Me.txt_tipo_3) Then
            If IsNumeric(Me.txt_tipo_4) Then
               rs.Open "SELECT * FROM VIA_FACTURAS WHERE foltda_codigo=" + Me.txt_tipo_1 + " and folest_codigo=" + Me.txt_tipo_2 + " and foldoc_codigo=" + Me.txt_tipo_3 + " and folconsecutivo=" + Me.txt_tipo_4, cnn_compucaja, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.lv_detalle.ListItems.Clear
                  Me.txt_total = Format(IIf(IsNull(rs!f_importetotal), 0, rs!f_importetotal), "###,###,##0.00")
                  Me.txt_rfc = IIf(IsNull(rs!cli_registrotributario), "", rs!cli_registrotributario)
                  Me.txt_nombre = IIf(IsNull(rs!cli_nombre), "", rs!cli_nombre)
                  Me.txt_cp = IIf(IsNull(rs!cli_cp), "", rs!cli_cp)
                  Me.txt_direccion = IIf(IsNull(rs!cli_domicilio), "", rs!cli_domicilio)
                  Me.txt_colonia = ""
                  Me.txt_pais = IIf(IsNull(rs!cli_pais), "", rs!cli_pais)
                  Me.txt_estado = IIf(IsNull(rs!cli_estado), "", rs!cli_estado)
                  Me.txt_municipio = IIf(IsNull(rs!cli_municipio), "", rs!cli_domicilio)
                  Me.txt_ciudad = IIf(IsNull(rs!cli_asentamiento), "", rs!cli_asentamiento)
                  Me.txt_correo = ""
                  While Not rs.EOF
                        Set list_item = Me.lv_detalle.ListItems.Add(, , rs!art_codigo)
                        list_item.SubItems(1) = IIf(IsNull(rs!art_codigo), "", rs!art_descripcion)
                        list_item.SubItems(2) = Format(IIf(IsNull(rs!art_descripcion), "", rs!dt_cantidad), "###,###,##0.00")
                        var_precio = IIf(IsNull(rs!dt_preciounitario), 0, rs!dt_preciounitario)
                        var_descuento = IIf(IsNull(rs!ddt_porcentaje), 0, rs!ddt_porcentaje)
                        var_precio2 = var_precio * (1 + (var_descuento / 100))
                        var_importe_descuento = var_precio2 - var_precio
                        list_item.SubItems(3) = Format(var_precio2, "###,###,##0.00")
                        list_item.SubItems(4) = Format(IIf(IsNull(rs!ddt_porcentaje), 0, rs!ddt_porcentaje), "###,###,##0.00")
                        list_item.SubItems(5) = Format(IIf(IsNull(var_importe_descuento), 0, var_importe_descuento), "###,###,##0.00")
                        list_item.SubItems(6) = Format(IIf(IsNull(var_precio), 0, var_precio), "###,###,##0.00")
                        list_item.SubItems(7) = Format(IIf(IsNull(rs!dt_precioventa), 0, rs!dt_precioventa), "###,###,##0.00")
                        rs.MoveNext
                  Wend
               Else
                  MsgBox "Folio de factura incorrecto", vbOKOnly, "ATENCION"
                  Me.lv_detalle.ListItems.Clear
                  Me.txt_rfc = ""
                  Me.txt_nombre = ""
                  Me.txt_direccion = ""
                  Me.txt_cp = ""
                  Me.txt_ciudad = ""
                  Me.txt_colonia = ""
                  Me.txt_estado = ""
                  Me.txt_municipio = ""
                  Me.txt_pais = ""
                  Me.txt_correo = ""
                  Me.txt_total = ""
               End If
               rs.Close
            Else
               MsgBox "Parametro 4 incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Parametro 3 incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Parametro 2 incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Paramentro 1 incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmd_factura_electronica_Click()
   var_ruta_documentos_electronicos = "\\facelectronica\fefiles\Conectorcan\envio\por_enviar"
   'var_ruta_documentos_electronicos = "C:"
   var_posible_factura_electronica = 0
   If Me.chk_factura_del_dia = 1 Then
      var_si = MsgBox("Se a seleccionado la facturación del dia ¿Desea continuar?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la factura del dia", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_posible_factura_electronica = 1
         Else
            var_posible_factura_electronica = 0
         End If
      Else
         var_posible_factura_electronica = 0
      End If
   Else
      var_posible_factura_electronica = 1
   End If
   If var_posible_factura_electronica = 1 Then
      If IsNumeric(Me.txt_total) Then
         rsaux10.Open "select * from TB_FACTURAS_COMPUCAJA_SID where vcha_FOL_folio_compucaja = '" + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux10.EOF Then
            cnn.BeginTrans
            rsaux1.Open "SELECT * FROM TB_sERIES WHERE VCHA_UOR_UNIDAD_ID = 'CANTIA'", cnn, adOpenDynamic, adLockOptimistic
            var_serie = IIf(IsNull(rsaux1!vcha_Ser_Serie_id), "", rsaux1!vcha_Ser_Serie_id)
            var_numero_factura = IIf(IsNull(rsaux1!inte_ser_factura), 0, rsaux1!inte_ser_factura)
            rsaux1.Close
            Open (App.Path & "\renombra" + Trim(Str(var_numero_factura)) + ".bat") For Output As #2
            If var_empresa = "31" Then
               Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".fi " + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".ff"
            End If
            Close #2
            rsaux3.Open "SELECT * FROM VIA_FACTURAS WHERE foltda_codigo=" + Me.txt_tipo_1 + " and folest_codigo=" + Me.txt_tipo_2 + " and foldoc_codigo=" + Me.txt_tipo_3 + " and folconsecutivo=" + Me.txt_tipo_4, cnn_compucaja, adOpenDynamic, adLockOptimistic
            var_subimporte = 0
            var_importe_Descuentos = 0
            
            var_rfc_cliente_1 = IIf(IsNull(rsaux3!cli_registrotributario), "", rsaux3!cli_registrotributario)
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
            
            If var_rfc_cliente = "XAXX010101000" Then
               var_desglose = 1
            Else
               var_desglose = 1.16
            End If
            
            
            While Not rsaux3.EOF
                  var_precio = IIf(IsNull(rsaux3!dt_preciounitario), 0, rsaux3!dt_preciounitario)
                  var_descuento_1 = IIf(IsNull(rsaux3!ddt_porcentaje), 0, rsaux3!ddt_porcentaje)
                  var_descuento_2 = 0
                  var_descuento_3 = 0
                  var_porcentaje = (100 - var_descuento_1) / 100
                  If var_porcentaje = 0 Then
                     var_precio = 0
                  Else
                     var_precio = (var_precio / var_porcentaje) / var_desglose
                  End If
                  var_importe_descuento_1_2 = (var_precio - (IIf(IsNull(rsaux3!dt_preciounitario), 0, rsaux3!dt_preciounitario) / var_desglose))
                  var_importe_Descuentos = var_importe_Descuentos + (var_importe_descuento_1_2 * rsaux3!dt_cantidad)
                  rsaux3.MoveNext
            Wend
            rsaux3.MoveFirst
            var_subimporte = (rsaux3!f_importetotal / var_desglose) + var_importe_Descuentos
            
            
         
         
         
         
         
         
         
         
         
         
            Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".fi") For Output As #1
            var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + var_serie + Chr(13) + "folio=" + CStr(var_numero_factura) + Chr(13)
            var_año = CStr(Year(rsaux3!F_FECHA))
            var_mes = CStr(Month(rsaux3!F_FECHA))
            var_dia = CStr(Day(rsaux3!F_FECHA))
            var_hora = CStr(Hour(rsaux3!F_FECHA))
            var_minuto = CStr(Minute(rsaux3!F_FECHA))
            var_segundo = CStr(Second(rsaux3!F_FECHA))
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
            var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
            var_cadena = var_cadena + "noAprobacion=" + Chr(13)
            var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
            var_cadena = var_cadena + "tipoDeComprobante=FACTURA" + Chr(13)
            var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
            var_cadena = var_cadena + "condicionesDePago=CONTADO" + Chr(13)
            var_total = rsaux3!f_importetotal
            Call numero_letras(rsaux3!f_importetotal, "1")
            var_cantidad_letra = canstr
            var_importe_iva = var_total - (var_total / var_desglose)
            var_subtotal = var_subimporte
            'MsgBox Format(CStr(var_subtotal), "###,###,##0.0000")
            var_cadena = var_cadena + "subtotal=" + Format(CStr(var_subtotal), "###,###,##0.0000") + Chr(13)
            'rsaux1.Open "select sum(ddt_importe) from via_facturas where foltda_codigo=" + Me.txt_tipo_1 + " and folest_codigo=" + Me.txt_tipo_2 + " and foldoc_codigo=" + Me.txt_tipo_3 + " and folconsecutivo=" + Me.txt_tipo_4, cnn_compucaja, adOpenDynamic, adLockOptimistic
            'If Not rsaux1.EOF Then
            '   var_importe_descuento = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
            'MsgBox var_cadena
               var_cadena = var_cadena + "descuento=" + Format(CStr(var_importe_Descuentos), "###,###,##0.0000") + Chr(13)
            'Else
            '   var_cadena = var_cadena + "descuento=0" + Chr(13)
            'End If
            'rsaux1.Close
            var_cadena = var_cadena + "descuento1=" + Chr(13)
            var_cadena = var_cadena + "descuento2=" + Chr(13)
            var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
            var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
            var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
            var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            rsaux1.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            var_certificado = rsaux1!vcha_emp_certificado
            var_expedido = rsaux1!vcha_emp_expedido
            var_cadena = var_cadena + "iva=" + Format(CStr(var_importe_iva), "###,###,##0.0000") + Chr(13)
            var_cadena = var_cadena + "total=" + Format(CStr(var_total), "###,###,##0.0000") + Chr(13)
            var_cadena = var_cadena + "retencion=" + Chr(13)
            var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
            var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Emisor>" + Chr(13)
            var_cadena = var_cadena + "erfc=" + rsaux1!VCHA_eMP_RFC + Chr(13)
            var_cadena = var_cadena + "enombre=" + rsaux1!VCHA_EMP_NOMBRE + Chr(13)
            var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
            var_cadena = var_cadena + "ecalle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
            var_cadena = var_cadena + "enoExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
            var_cadena = var_cadena + "enoInterior=" + Chr(13)
            var_cadena = var_cadena + "ecolonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
            var_cadena = var_cadena + "elocalidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
            var_cadena = var_cadena + "ereferencia=" + Chr(13)
            var_cadena = var_cadena + "emunicipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
            var_cadena = var_cadena + "eestado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
            var_cadena = var_cadena + "epais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
            var_cadena = var_cadena + "ecodigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
            var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux1!VCHA_EMP_TELEFONO), "", rsaux1!VCHA_EMP_TELEFONO) + Chr(13)
            var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL) + Chr(13)
            correo = IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL)
            var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
            
            var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "ex_calle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
            var_cadena = var_cadena + "ex_noExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
            var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
            var_cadena = var_cadena + "ex_colonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
            var_cadena = var_cadena + "ex_localidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
            var_cadena = var_cadena + "ex_referencia=" + Chr(13)
            var_cadena = var_cadena + "ex_municipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
            var_cadena = var_cadena + "ex_estado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
            var_cadena = var_cadena + "ex_pais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
            var_cadena = var_cadena + "ex_codigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
            var_cadena = var_cadena + "</ExpedidoEn>"
            
            
            
            
            var_cadena = var_cadena + "<Receptor>" + Chr(13)
            var_cadena = var_cadena + "noCliente=" + rsaux3!f_cliente + Chr(13)
                                            
            var_rfc_cliente_1 = IIf(IsNull(rsaux3!cli_registrotributario), "", rsaux3!cli_registrotributario)
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
            var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
            var_cadena = var_cadena + "nombre=" + rsaux3!cli_nombre + " " + IIf(IsNull(rsaux3!Cli_ApellidoPaterno), "", rsaux3!Cli_ApellidoPaterno) + " " + IIf(IsNull(Cli_Apellidomaterno), "", rsaux3!Cli_Apellidomaterno) + Chr(13)
            var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Cliente>" + Chr(13)
            var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rsaux3!cli_domicilio), "", rsaux3!cli_domicilio) + Chr(13)
            var_cadena = var_cadena + "calle=" + Chr(13)
            var_cadena = var_cadena + "noExterior=" + Chr(13)
            var_cadena = var_cadena + "noInterior=" + Chr(13)
            var_cadena = var_cadena + "colonia=" + Chr(13)
            var_cadena = var_cadena + "localidad=" + IIf(IsNull(rsaux3!cli_asentamiento), "", rsaux3!cli_asentamiento) + Chr(13)
            rsaux1.Close
            var_cadena = var_cadena + "referencia=" + Chr(13)
            var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux3!cli_municipio), "", rsaux3!cli_municipio) + Chr(13)
            var_cadena = var_cadena + "estado=" + IIf(IsNull(rsaux3!cli_estado), "", rsaux3!cli_estado) + Chr(13)
            VAR_NOMBRE_PAIS = IIf(IsNull(rsaux3!cli_pais), "MEXICO", rsaux3!cli_pais)
            If Trim(VAR_NOMBRE_PAIS) = "" Then
               VAR_NOMBRE_PAIS = "MEXICO"
            End If
            var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
            var_cadena = var_cadena + Chr(13)
            var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rsaux3!cli_cp), "", rsaux3!cli_cp) + Chr(13)
            var_cadena = var_cadena + "tel=" + Chr(13)
            var_cadena = var_cadena + "email=" + Me.txt_correo + Chr(13)
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

            var_k = 0
            If Me.chk_factura_del_dia = 1 Then
               var_k = var_k + 1
               pxx = CStr(var_k)
               If Len(pxx) = 1 Then
                  pxx = "0" + pxx
               End If
               var_cadena = var_cadena + "p" + pxx + "_cantidad=1" + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_unidad=PZA" + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=VTAPG" + Chr(13)
               var_linea = "VENTA AL PUBLICO EN GENERAL"
               var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_subimporte), "###,###,##0.00") + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_subimporte), "###,###,##0.00") + Chr(13)
            Else
               While Not rsaux3.EOF
                     var_k = var_k + 1
                     pxx = CStr(var_k)
                     If Len(pxx) = 1 Then
                        pxx = "0" + pxx
                     End If
                     var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rsaux3!dt_cantidad), 0, rsaux3!dt_cantidad)) + Chr(13)
                     var_cadena = var_cadena + "p" + pxx + "_unidad=PZA" + Chr(13)
                     var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + IIf(IsNull(rsaux3!art_codigo), "", rsaux3!art_codigo) + Chr(13)
                     var_linea = IIf(IsNull(rsaux3!art_codigo), "", rsaux3!art_codigo) + " " + IIf(IsNull(rsaux3!art_descripcion), "", rsaux3!art_descripcion)
                     var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                     var_precio = IIf(IsNull(rsaux3!dt_preciounitario), 0, rsaux3!dt_preciounitario)
                     var_descuento_1 = IIf(IsNull(rsaux3!ddt_porcentaje), 0, rsaux3!ddt_porcentaje)
                     var_descuento_2 = 0
                     var_descuento_3 = 0
                     var_porcentaje = (100 - var_descuento_1) / 100
                     If var_porcentaje = 0 Then
                        var_precio = 0
                     Else
                        var_precio = (var_precio / var_porcentaje) / var_desglose
                     End If
                     var_importe_descuento_1_2 = (IIf(IsNull(rsaux3!dt_preciounitario), 0, rsaux3!dt_preciounitario) - var_precio)
                     var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rsaux3!dt_preciounitario), 0, rsaux3!dt_preciounitario) - var_precio)
                     var_precio = var_precio * ((100 - var_descuento_2) / 100)
                     var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rsaux3!dt_preciounitario), 0, rsaux3!dt_preciounitario) - (var_importe_descuento_1_2 + var_precio))
                     var_precio = var_precio * ((100 - var_descuento_3) / 100)
                     var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_precio), "###,###,##0.00") + Chr(13)
                     var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_precio * CStr(IIf(IsNull(rsaux3!dt_cantidad), 0, rsaux3!dt_cantidad))), "###,###,##0.00") + Chr(13)
                                                  
                     rsaux3.MoveNext
               Wend
               rsaux3.MoveFirst
            End If
            var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Otros>" + Chr(13)
            var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
            var_cadena = var_cadena + "cant_letra=" + var_cantidad_letra + Chr(13)
            var_cadena = var_cadena + "factoriva=16%" + Chr(13)
            var_cadena = var_cadena + "moneda=PESOS" + Chr(13)
            var_cadena = var_cadena + "tipodeCambio=1" + Chr(13)
            var_cadena = var_cadena + "pedido=" + Chr(13)
            var_cadena = var_cadena + "Embarque=" + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + Chr(13)
            var_referencia_Bancaria = ""
            var_referencia_Bancaria = ""
            var_cadena = var_cadena + "referenciabancaria=" + var_referencia_Bancaria + Chr(13)
            var_cadena = var_cadena + "fechaPedido=" + var_cadena_fecha + Chr(13)
            var_cadena = var_cadena + "expedicion=" + var_expedido + Chr(13)
            var_cadena = var_cadena + "observaciones=" + Chr(13)
            var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
            var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
            var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
            var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
            var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
            var_cadena = var_cadena + "agente=" + CStr(IIf(IsNull(rsaux3!f_vendedor), "", rsaux3!f_vendedor)) + Chr(13)
            If var_empresa = "15" Then
               var_cadena = var_cadena + "formato=MHESTAMPADOS_V01.DAT" + Chr(13)
            End If
            If var_empresa = "16" Then
               var_cadena = var_cadena + "formato=MHMULTIBONDEADOS_V01.DAT" + Chr(13)
            End If
            If var_empresa = "31" Then
               var_cadena = var_cadena + "formato=MHCANTIA_V01.DAT" + Chr(13)
            End If
                              
            var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<addenda>" + Chr(13)
            var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "</Factura>"
            Print #1, var_cadena
            Close #1
            var_Archivo = App.Path & "\renombra" + Trim(Str(var_numero_factura)) + ".bat"
            x = Shell(var_Archivo, vbHide)
            rsaux1.Open "INSERT INTO TB_FACTURAS_COMPUCAJA_SID (VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_FOL_FOLIO_COMPUCAJA) VALUES ('" + var_serie + "', '" + CStr(var_numero_factura) + "','" + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux1.Open "UPDATE TB_sERIES SET INTE_SER_FACTURA = INTE_SER_FACTURA + 1 WHERE VCHA_UOR_UNIDAD_ID = 'CANTIA'", cnn, adOpenDynamic, adLockOptimistic
            rsaux3.Close
            cnn.CommitTrans
            Me.chk_factura_del_dia.Value = 0
           
            'var_ruta_documentos_electronicos = "\\facelectronica\fefiles\ConectorVIA\envio\enviados"
            'Open (App.Path & "\EJECTUAPDF.bat") For Output As #2
            'If var_empresa = "31" Then
            '   Print #2, "START " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".PDF"
            'End If
            'Close #2
            '
            'var_Archivo = App.Path & "\EJECTUAPDF.bat"
            'x = Shell(var_Archivo, vbHide)
            MsgBox "Se a enviado la factura electronica", vbOKOnly, "ATENCION"
         Else
            MsgBox "El folio " + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + ", se encuentra en la factura " + CStr(rsaux10!inte_Car_numero), vbOKOnly, "ATENCION"
         End If
         rsaux10.Close
      Else
         MsgBox "Factura incorrecta", vbOKOnly, "ATENCION"
         Me.chk_factura_del_dia = 0
      End If
   Else
      MsgBox "Se a cancelado la facturación electrónica", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_folio_Change()
   Me.lv_detalle.ListItems.Clear
   Me.txt_rfc = ""
   Me.txt_nombre = ""
   Me.txt_direccion = ""
   Me.txt_cp = ""
   Me.txt_ciudad = ""
   Me.txt_colonia = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_pais = ""
   Me.txt_correo = ""
   Me.txt_total = ""
End Sub

Private Sub txt_tipo_1_Change()
   Me.lv_detalle.ListItems.Clear
   Me.txt_rfc = ""
   Me.txt_nombre = ""
   Me.txt_direccion = ""
   Me.txt_cp = ""
   Me.txt_ciudad = ""
   Me.txt_colonia = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_pais = ""
   Me.txt_correo = ""
   Me.txt_total = ""
End Sub

Private Sub txt_tipo_1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 39 Then
      Me.txt_tipo_2.SetFocus
   End If
End Sub

Private Sub txt_tipo_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_2.SetFocus
   End If
End Sub

Private Sub txt_tipo_2_Change()
   Me.lv_detalle.ListItems.Clear
   Me.txt_rfc = ""
   Me.txt_nombre = ""
   Me.txt_direccion = ""
   Me.txt_cp = ""
   Me.txt_ciudad = ""
   Me.txt_colonia = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_pais = ""
   Me.txt_correo = ""
   Me.txt_total = ""
End Sub

Private Sub txt_tipo_2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_tipo_1.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_tipo_3.SetFocus
   End If
End Sub

Private Sub txt_tipo_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_3.SetFocus
   End If
End Sub

Private Sub txt_tipo_3_Change()
   Me.lv_detalle.ListItems.Clear
   Me.txt_rfc = ""
   Me.txt_nombre = ""
   Me.txt_direccion = ""
   Me.txt_cp = ""
   Me.txt_ciudad = ""
   Me.txt_colonia = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_pais = ""
   Me.txt_correo = ""
   Me.txt_total = ""
End Sub

Private Sub txt_tipo_3_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_tipo_2.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_tipo_4.SetFocus
   End If
End Sub

Private Sub txt_tipo_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_4.SetFocus
   End If
End Sub

Private Sub txt_tipo_4_Change()
   Me.lv_detalle.ListItems.Clear
   Me.txt_rfc = ""
   Me.txt_nombre = ""
   Me.txt_direccion = ""
   Me.txt_cp = ""
   Me.txt_ciudad = ""
   Me.txt_colonia = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_pais = ""
   Me.txt_correo = ""
   Me.txt_total = ""
End Sub

Private Sub txt_tipo_4_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_tipo_3.SetFocus
   End If
   If KeyCode = 39 Then
      Me.cmd_buscar.SetFocus
   End If
End Sub

Private Sub txt_tipo_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_buscar.SetFocus
   End If
End Sub

Private Sub txt_total_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
End Sub
