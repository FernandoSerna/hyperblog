VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcancela_facturas_devolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de facturas para refacturación"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   10350
   Begin VB.Frame frm_cliente 
      Height          =   3525
      Left            =   45
      TabIndex        =   50
      Top             =   1035
      Width           =   10110
      Begin VB.Frame frm_lista 
         Height          =   2400
         Left            =   1740
         TabIndex        =   64
         Top             =   825
         Width           =   5685
         Begin MSComctlLib.ListView lv_lista 
            Height          =   1875
            Left            =   30
            TabIndex        =   65
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
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   7408
            EndProperty
         End
         Begin VB.Label lbl_lista 
            BackColor       =   &H8000000D&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   30
            TabIndex        =   66
            Top             =   120
            Width           =   5610
         End
      End
      Begin VB.TextBox txt_nombre_establecimiento_nuevo 
         Height          =   315
         Left            =   2745
         TabIndex        =   34
         Top             =   3060
         Width           =   6660
      End
      Begin VB.TextBox txt_nombre_cliente_nuevo 
         Height          =   315
         Left            =   2220
         TabIndex        =   23
         Top             =   915
         Width           =   7185
      End
      Begin VB.TextBox txt_clave_establecimiento_nuevo 
         Height          =   315
         Left            =   1485
         TabIndex        =   33
         Top             =   3060
         Width           =   1245
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frmcancela_facturas_devolucion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancelar"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Picture         =   "frmcancela_facturas_devolucion.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Aceptar"
         Top             =   375
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   30
         TabIndex        =   62
         Top             =   645
         Width           =   10035
      End
      Begin VB.TextBox txt_colonia_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5580
         TabIndex        =   25
         Top             =   1260
         Width           =   3840
      End
      Begin VB.TextBox txt_rfc_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5580
         TabIndex        =   29
         Top             =   1950
         Width           =   1770
      End
      Begin VB.TextBox txt_cp_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   30
         Top             =   1950
         Width           =   1575
      End
      Begin VB.TextBox txt_pais_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   28
         Top             =   1950
         Width           =   3870
      End
      Begin VB.TextBox txt_estado_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5595
         TabIndex        =   27
         Top             =   1605
         Width           =   3855
      End
      Begin VB.TextBox txt_ciudad_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Top             =   1605
         Width           =   3855
      End
      Begin VB.TextBox txt_domicilio_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Top             =   1260
         Width           =   3840
      End
      Begin VB.TextBox txt_clave_cliente_nuevo 
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   930
         Width           =   1245
      End
      Begin VB.TextBox txt_descuento_2_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   32
         Top             =   2625
         Width           =   1230
      End
      Begin VB.TextBox txt_descuento_1_nuevo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   31
         Top             =   2280
         Width           =   1230
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   240
         TabIndex        =   63
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del cliente a refacturar"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   45
         TabIndex        =   61
         Top             =   120
         Width           =   10020
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Left            =   4920
         TabIndex        =   60
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Left            =   7515
         TabIndex        =   59
         Top             =   2010
         Width           =   255
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   4920
         TabIndex        =   57
         Top             =   1665
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   4920
         TabIndex        =   56
         Top             =   2010
         Width           =   360
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   225
         TabIndex        =   55
         Top             =   1665
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   225
         TabIndex        =   54
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   225
         TabIndex        =   53
         Top             =   975
         Width           =   525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Descuento Pago Correcto:"
         Height          =   195
         Left            =   735
         TabIndex        =   52
         Top             =   2760
         Width           =   1890
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Descuento Pronto Pago:"
         Height          =   195
         Left            =   765
         TabIndex        =   51
         Top             =   2340
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      Height          =   2835
      Left            =   105
      TabIndex        =   0
      Top             =   1395
      Width           =   10110
      Begin VB.TextBox txt_clave_establecimiento 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   2370
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   2370
         Width           =   6195
      End
      Begin VB.TextBox txt_clave_titular 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   630
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   630
         Width           =   6195
      End
      Begin VB.TextBox txt_clave_agente 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   270
         Width           =   6195
      End
      Begin VB.TextBox txt_colonia 
         Height          =   315
         Left            =   6045
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1335
         Width           =   3840
      End
      Begin VB.TextBox txt_rfc 
         Height          =   315
         Left            =   6045
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2025
         Width           =   1770
      End
      Begin VB.TextBox txt_cp 
         Height          =   315
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2025
         Width           =   1575
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2025
         Width           =   3870
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   6045
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txt_ciudad 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txt_domicilio 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1335
         Width           =   3840
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   6195
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   225
         TabIndex        =   78
         Top             =   2430
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   225
         TabIndex        =   75
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Left            =   5385
         TabIndex        =   49
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Left            =   7980
         TabIndex        =   42
         Top             =   2085
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   2085
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   5385
         TabIndex        =   40
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   5385
         TabIndex        =   39
         Top             =   2085
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   225
         TabIndex        =   38
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   225
         TabIndex        =   37
         Top             =   1395
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   225
         TabIndex        =   36
         Top             =   1050
         Width           =   525
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Factura "
      Height          =   900
      Left            =   105
      TabIndex        =   67
      Top             =   435
      Width           =   10110
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   3060
         TabIndex        =   1
         Top             =   330
         Width           =   1230
      End
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   5550
         TabIndex        =   2
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   2325
         TabIndex        =   69
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   4815
         TabIndex        =   68
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcancela_facturas_devolucion.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmcancela_facturas_devolucion.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Refacturar"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9810
      Picture         =   "frmcancela_facturas_devolucion.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Importes "
      Height          =   1365
      Left            =   105
      TabIndex        =   43
      Top             =   4320
      Width           =   10110
      Begin VB.TextBox txt_iva 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7710
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7710
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   1485
      End
      Begin VB.TextBox txt_descuento_2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   750
         Width           =   1230
      End
      Begin VB.TextBox txt_subimporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7710
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   255
         Width           =   1485
      End
      Begin VB.TextBox txt_descuento_1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   405
         Width           =   1230
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Importe Neto:"
         Height          =   195
         Left            =   6150
         TabIndex        =   48
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Left            =   6150
         TabIndex        =   47
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Subimporte:"
         Height          =   195
         Left            =   6150
         TabIndex        =   46
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descuento Pago Correcto:"
         Height          =   195
         Left            =   810
         TabIndex        =   45
         Top             =   810
         Width           =   1890
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descuento Pronto Pago:"
         Height          =   195
         Left            =   810
         TabIndex        =   44
         Top             =   465
         Width           =   1755
      End
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   105
      TabIndex        =   35
      Top             =   270
      Width           =   10110
   End
End
Attribute VB_Name = "frmcancela_facturas_devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_serie As String
Dim var_agente As String
Dim txt_numero As String
Dim txt_movimiento As String
Dim var_tipo_lista As Integer

Private Sub cmd_aceptar_Click()
Dim si As Integer
Dim var_almacen_calidad As String
Dim var_movimiento_calidad As String
Dim var_movimiento_traspaso As String
Dim var_almacen_surtido As String
Dim var_clave_Causa_devolucion As Integer
Dim var_primera_vez As Boolean
   If Trim(txt_clave_cliente_nuevo) <> "" Then
      If Trim(txt_clave_establecimiento_nuevo) <> "" Then
         si = MsgBox("¿Desea hacer la reimpresión de la factura " + txt_factura, vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la reimpresión de la factura " + txt_factura, vbYesNo, "ATENCION")
            If si = 6 Then
               rs.Open "SELECT VCHA_REF_ALMACEN_CALIDAD, VCHA_REF_MOVIMIENTO_CALIDAD, VCHA_REF_ALMACEN_SURTIDO, VCHA_REF_MOVIMIENTO_TRASPASO FROM TB_DATOS_REFACTURACION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_almacen_calidad = IIf(IsNull(rs!VCHA_REF_ALMACEN_CALIDAD), "", rs!VCHA_REF_ALMACEN_CALIDAD)
                  If var_almacen_calidad <> "" Then
                     var_movimiento_calidad = IIf(IsNull(rs!VCHA_REF_MOVIMIENTO_CALIDAD), "", rs!VCHA_REF_MOVIMIENTO_CALIDAD)
                     If var_movimiento_calidad <> "" Then
                        var_almacen_surtido = IIf(IsNull(rs!VCHA_REF_ALMACEN_SURTIDO), "", rs!VCHA_REF_ALMACEN_SURTIDO)
                        If var_almacen_surtido <> "" Then
                           var_movimiento_traspaso = IIf(IsNull(rs!VCHA_REF_MOVIMIENTO_TRASPASO), "", rs!VCHA_REF_MOVIMIENTO_TRASPASO)
                           If var_movimiento_traspaso <> "" Then
                              rsaux.Open "SELECT INTE_CDE_CAUSA_ID FROM TB_CAUSAS_DEVOLUCION WHERE INTE_CDE_REFACTURACION = 1", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_clave_Causa_devolucion = IIf(IsNull(rsaux!INTE_CDE_CAUSA_ID), 0, rsaux!INTE_CDE_CAUSA_ID)
                                 If var_clave_Causa_devolucion <> 0 Then
                                    var_cadena = "SELECT     SUM(FLOA_SAL_CANTIDAD) AS floa_sal_cantidad, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID,"
                                    var_cadena = var_cadena + " VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO, VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE , DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                    var_cadena = var_cadena + " From dbo.TB_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Car_documento = 'FA' and vcha_Ser_Serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_factura + " GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO,"
                                    var_cadena = var_cadena + " VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE, DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                    rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    'rsaux1.Open "select * from tb_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Car_documento = 'FA' and vcha_Ser_Serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
                                    var_primera_vez = True
                                    Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
                                    Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
                                    Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
                                    Set TB_BLOQUEOS = New TB_BLOQUEOS
                                    Dim var_inserta As Boolean
                                    Dim var_factura As Integer
                                    Dim var_numero_folio As Double
                                    Dim var_cantidad_leida As Double
                                    Dim var_numero_folio_calidad As Double
                                    While Not rsaux1.EOF
                                          If Trim(rsaux1!VCHA_ART_ARTICULO_ID) <> "" Then
                                             bandera_suma = False
                                             If var_primera_vez = True Then
                                                var_inserta = False
                                                If rsaux9.State = 1 Then
                                                   rsaux9.Close
                                                End If
                                                rsaux9.Open "select vcha_mon_moneda_id from tb_clientes where vcha_cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                                var_clave_moneda = IIf(IsNull(rsaux9!vcha_mon_moneda_id), "", rsaux9!vcha_mon_moneda_id)
                                                rsaux9.Close
                                                var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_calidad, var_movimiento_calidad, Now, CDbl(var_numero_folio), 0, CStr(Me.txt_clave_cliente), "", "", CStr(var_almacen_calidad), "", CStr(var_clave_usuario_global), fun_NombrePc, "", "", "Devolución por refacturación de la factura " + Me.txt_factura, CStr(txt_clave_establecimiento), "B", CStr(Me.txt_clave_titular), CStr(txt_clave_agente), 0, 0, 0, CStr(var_clave_moneda), 1)
                                                var_numero_folio = var_numero_folio_regreso
                                                var_numero_folio_calidad = var_numero_folio
                                                var_global_bloqueado = 1
                                                var_inserta = False
                                                var_inserta = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, "DEVOLUCION" + Trim(var_clave_movimiento) + Trim(Str(var_numero_folio)), Now, var_clave_usuario_global, fun_NombrePc)
                                                var_solo_lectura = False
                                                txt_folio = var_numero_folio
                                                var_primera_vez = False
                                             End If
       
                                             var_costo = IIf(IsNull(rsaux1!floa_Sal_costo), 0, rsaux1!floa_Sal_costo)
                                             var_precio = IIf(IsNull(rsaux1!floa_Sal_precio), 0, rsaux1!floa_Sal_precio)
                                             var_año = 2005
                                             Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_calidad + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_calidad + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + rsaux1!VCHA_ART_ARTICULO_ID + "' and vcha_emp_empresa_id = '" + var_empresa + "'"
                                             rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             var_cantidad_leida = IIf(IsNull(rsaux1!floa_Sal_Cantidad), 0, rsaux1!floa_Sal_Cantidad)
                                             If Not rsaux2.EOF Then
                                                var_inserta = False
                                                var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_calidad), CStr(var_movimiento_calidad), var_numero_folio, CStr(rsaux2!VCHA_ART_ARTICULO_ID), var_cantidad_leida, CDbl(var_año))
                                                rsaux2.Close
                                             Else
                                                var_inserta = False
                                                var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, CStr(var_almacen_calidad), CStr(var_movimiento_calidad), var_numero_folio, CStr(rsaux1!VCHA_ART_ARTICULO_ID), var_cantidad_leida, CDbl(var_costo), CDbl(var_precio), "0", "", CDbl(var_año))
                                                rsaux2.Close
                                             End If
                                          End If
                                          rsaux1.MoveNext
                                    Wend
                                    rsaux1.Close
                                 
                                 
                                    Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
                                    Set TB_ENTRADAS_I = New TB_ENTRADAS_I
                                    Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
                                    If var_numero_folio > 0 Then
                                      
                                       Set TB_DEVOLUCIONES_INSERTA = New TB_DEVOLUCIONES_INSERTA
                                       cnn.BeginTrans
                                       Cadena = "select * from tb_temporal_entradas where vcha_alm_almacen_id = '" + var_almacen_calidad + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_calidad + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "'"
                                       rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_inserta = False
                                          var_inserta = TB_ENTRADAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_calidad, var_movimiento_calidad, var_numero_folio, "", "", 0)
                                       End If
                                       rsaux2.Close
                                       var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_calidad, var_movimiento_calidad, var_numero_folio, "I", Now, 1)
                                       cnn.CommitTrans
                                    End If
                                    
                                    rsaux3.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_cAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' and  INTE_CAR_NUMERO = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux3.EOF Then
                                       var_tipo_Cambio = IIf(IsNull(rsaux3!floa_car_tipo_cambio), 1, rsaux3!floa_car_tipo_cambio)
                                    Else
                                       var_tipo_Cambio = 1
                                    End If
                                    rsaux3.Close
                                    
                                    'MsgBox CStr(var_tipo_Cambio)
                                    var_cadena = "SELECT     SUM(FLOA_SAL_CANTIDAD) AS floa_sal_cantidad, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID,"
                                    var_cadena = var_cadena + " VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO, VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE , DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                    var_cadena = var_cadena + " From dbo.TB_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Car_documento = 'FA' and vcha_Ser_Serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_factura + " GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO,"
                                    var_cadena = var_cadena + " VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE, DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                    
                                    rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    'rsaux3.Open "select * from tb_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Car_documento = 'FA' and vcha_Ser_Serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
                                    
                                    
                                    var_factura_ceros = 0
                                    If Not rsaux3.EOF Then
                                       var_movimiento_salida = rsaux3!VCHA_MOV_MOVIMIENTO_ID
                                       rsaux4.Open "select inte_emo_Factura_ceros from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + rsaux3!VCHA_MOV_MOVIMIENTO_ID + "' and inte_emo_numero = " + CStr(rsaux3!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux4.EOF Then
                                          var_factura_ceros = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                       End If
                                       rsaux4.Close
                                    End If
                                    
                                    rsaux5.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux5.EOF Then
                                       var_iva = IIf(IsNull(rsaux5!FLOA_TPE_IVA), 0, rsaux5!FLOA_TPE_IVA)
                                    End If
                                    If var_iva = 0 Then
                                       If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "16" Or var_empresa = "15" Then
                                          var_iva = 16
                                       End If
                                    End If
                                    rsaux5.Close
                                    
                                    While Not rsaux3.EOF
                                          
                                          var_contador = rsaux3!floa_Sal_Cantidad
                                          var_cantidad_pasar = rsaux3!floa_Sal_Cantidad
                                          'aqui
                                          While var_contador > 0
                                                'If var_contador >= 1 Then
                                                '   var_cantidad_pasar = 1
                                                '   var_contador = var_contador - 1
                                                'Else
                                                '   var_cantidad_pasar = var_contador
                                                '   var_contador = 0
                                                'End If
                                                var_consecutivo = var_consecutivo + 1
                                                var_costo = IIf(IsNull(rsaux3!floa_Sal_costo), 0, rsaux3!floa_Sal_costo)
                                                var_precio = IIf(IsNull(rsaux3!floa_Sal_precio), 0, rsaux3!floa_Sal_precio)
                                                var_descuento_1 = IIf(IsNull(rsaux3!FLOA_SAL_DESCUENTO_1), 0, rsaux3!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux3!FLOA_SAL_DESCUENTO_2), 0, rsaux3!FLOA_SAL_DESCUENTO_2)
                                                var_descuento_3 = 0
                                                vcha_agr_agrupador_id = IIf(IsNull(rsaux3!vcha_agr_agrupador_id), "", rsaux3!vcha_agr_agrupador_id)
                                                vcha_dev_descripcion_agrupador = IIf(IsNull(rsaux3!vcha_sal_descripcion_factura), "", rsaux3!vcha_sal_descripcion_factura)
                                                'var_inserta = TB_DEVOLUCIONES_INSERTA.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen_calidad), CStr(var_movimiento_calidad), CDbl(var_numero_folio), CStr(rsaux3!VCHA_aRT_aRTICULO_ID), 0, 0, "", CInt(var_consecutivo), "", CDbl(var_costo), CDbl(var_precio), CDbl(var_descuento_1), CDbl(var_descuento_2), CDbl(var_descuento_3), CDbl(var_iva), CDbl(Me.txt_factura), CStr("DEVOLUCION POR REFACTURACION DE LA FACTURA " + Me.txt_factura), CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(Me.txt_serie), CInt(2005))
                                                
                                                'rsaux5.Open "update tb_Devoluciones set CHAR_CDE_ESTATUS = 'I', VCHA_FAG_FAMILIA_AGRUPADOR_ID = '', VCHA_AGR_AGRUPADOR_ID = '" + vcha_agr_agrupador_id + "', VCHA_DEV_DESCRIPCION_AGRUPADOR ='" + VCHA_DEV_DESCRIPCION_AGRUPADOR + "' , floa_dev_cantidad = " + CStr(var_cantidad_pasar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_calidad + "' and vcha_mov_movimiento_id = '" + var_movimiento_calidad + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_art_articulo_id = '" + rsaux3!VCHA_aRT_aRTICULO_ID + "' and inte_cde_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                                var_cadena = "insert into tb_Devoluciones   (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID,                   VCHA_ALM_ALMACEN_ID,        VCHA_MOV_MOVIMIENTO_ID,               INTE_EMO_NUMERO,                VCHA_ART_ARTICULO_ID,            CHAR_CDE_ESTATUS, INTE_CDE_CONSECUTIVO,   VCHA_CDE_DESTINO,   FLOA_CDE_COSTO,        FLOA_CDE_PRECIO,             FLOA_CDE_DESCUENTO_1,              FLOA_CDE_DESCUENTO_2, FLOA_CDE_DESCUENTO_3,               FLOA_CDE_IVA,        INTE_FAC_FACTURA,          VCHA_CDE_REFERENCIA,                                                         VCHA_MON_MONEDA_ID,           FLOA_DEV_TIPO_CAMBIO,               VCHA_SER_SERIE_ID, INTE_DEV_AÑO,  VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID,            VCHA_DEV_DESCRIPCION_AGRUPADOR, floa_dev_cantidad )"
                                                var_cadena = var_cadena + "      values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_calidad + "', '" + var_movimiento_calidad + "', " + CStr(var_numero_folio) + ", '" + CStr(rsaux3!VCHA_ART_ARTICULO_ID) + "','I', " + CStr(CInt(var_consecutivo)) + ", '', " + CStr(CDbl(var_costo)) + ", " + CStr(var_precio) + ", " + CStr(var_descuento_1) + ", " + CStr(var_descuento_2) + ", " + CStr(var_descuento_3) + ", " + CStr(var_iva) + ", " + Me.txt_factura + ", '" + CStr("DEVOLUCION POR REFACTURACION DE LA FACTURA " + Me.txt_factura) + "', '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + Me.txt_serie + "', 2005,               '',                      '" + vcha_agr_agrupador_id + "','" + Mid(vcha_dev_descripcion_agrupador, 1, 50) + "'," + CStr(var_cantidad_pasar) + ")"
                                                'MsgBox var_cadena
                                                rsaux5.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                                
                                                var_cadena = "INSERT INTO TB_DETALLE_DEVOLUCION_REAL (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, INTE_CDE_CONSECUTIVO, INTE_CDE_CAUSA_ID, INTE_CDE_NOTA_CREDITO) VALUES"
                                                var_cadena = var_cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_calidad + "', '" + var_movimiento_calidad + "', " + CStr(var_numero_folio) + ",  '" + rsaux3!VCHA_ART_ARTICULO_ID + "',  " + CStr(var_consecutivo) + ", " + CStr(var_clave_Causa_devolucion) + ", 1)"
                                                rsaux5.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                                var_contador = 0
                                          Wend
                                          rsaux3.MoveNext
                                    Wend
                                    rsaux3.Close
                                    
                                    
                                    cnn.BeginTrans
                                                                                                          
                                                                       
                                                                       

                                    rsaux3.Open "SELECT INTE_EMO_NUMERO from TB_FOLIOS_MOVIMIENTOS where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_traspaso + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux3.EOF Then
                                       rsaux4.Open "INSERT INTO TB_FOLIOS_MOVIMIENTOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_movimiento_traspaso + "',1)"
                                       VAR_FOLIO_TRASPASO = 0
                                    Else
                                       rsaux4.Open "UPDATE TB_FOLIOS_MOVIMIENTOS SET INTE_EMO_NUMERO = isnull(INTE_EMO_NUMERO,0) + 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_traspaso + "'", cnn, adOpenDynamic, adLockOptimistic
                                       VAR_FOLIO_TRASPASO = IIf(IsNull(rsaux3!INTE_EMO_NUMERO), 0, rsaux3!INTE_EMO_NUMERO)
                                    End If
                                    VAR_FOLIO_TRASPASO = VAR_FOLIO_TRASPASO + 1
                                    rsaux3.Close
                                    cnn.CommitTrans
                                    var_cadena = "UPDATE TB_DEVOLUCIONES SET VCHA_CDE_DESTINO = '" + var_almacen_surtido + "', INTE_CDE_NUMERO_DESTINO = " + CStr(VAR_FOLIO_TRASPASO) + ", VCHA_CDE_MOVIMIENTO_DESTINO = '" + var_movimiento_traspaso + "' WHERE  (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_calidad + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_calidad + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio)
                                    rsaux3.Open "UPDATE TB_DEVOLUCIONES SET VCHA_CDE_DESTINO = '" + var_almacen_surtido + "', INTE_CDE_NUMERO_DESTINO = " + CStr(VAR_FOLIO_TRASPASO) + ", VCHA_CDE_MOVIMIENTO_DESTINO = '" + var_movimiento_traspaso + "' WHERE  VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_calidad + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_calidad + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   
                                    var_cadena = "INSERT INTO TB_ENCABEZADO_MOVIMIENTOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, DTIM_EMO_FECHA, INTE_EMO_NUMERO, INTE_EMO_NUMERO_ORIGEN, VCHA_CLI_CLAVE_ID, VCHA_PRO_PROVEEDOR_ID, VCHA_EMO_ALMACEN_ORIGEN, VCHA_EMO_ALMACEN_DESTINO, CHAR_EMO_ESTATUS, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_EMO_FACTURA, VCHA_EMO_MOVIMIENTO_ORIGEN, VCHA_EMO_REFERENCIA, VCHA_ESB_ESTABLECIMIENTO_ID,"
                                    var_cadena = var_cadena + " CHAR_EMO_BLOQUEADO, VCHA_TIT_TITULAR_ID, VCHA_AGE_AGENTE_ID, FLOA_EMO_DESCUENTO_1, FLOA_EMO_DESCUENTO_2, FLOA_EMO_DESCUENTO_3, VCHA_MON_MONEDA_ID, FLOA_EMO_TIPO_CAMBIO) VALUES"
                                    var_cadena = var_cadena + "('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_calidad + "', '" + var_movimiento_traspaso + "',  GETDATE(), '" + CStr(VAR_FOLIO_TRASPASO) + "', " + Me.txt_factura + ", '" + Me.txt_clave_cliente + "', '', '" + var_almacen_calidad + "', '" + var_almacen_surtido + "', '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '', '', 'REFACTURACION FACTURA " + Me.txt_factura + "', '" + Me.txt_clave_establecimiento + "', '', '" + Me.txt_clave_titular + "', '" + Me.txt_clave_agente + "', 0, 0, 0, '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ")"
                                    rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    
                                    var_cadena = "SELECT     SUM(FLOA_SAL_CANTIDAD) AS floa_sal_cantidad, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID,"
                                    var_cadena = var_cadena + " VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO, VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE , DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                    var_cadena = var_cadena + " From dbo.TB_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Car_documento = 'FA' and vcha_Ser_Serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_factura + " GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO,"
                                    var_cadena = var_cadena + " VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE, DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                    rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    'rsaux3.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND INTE_CAR_NUMERO = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux3.EOF
                                          var_costo = IIf(IsNull(rsaux3!floa_Sal_costo), 0, rsaux3!floa_Sal_costo)
                                          var_cantidad = IIf(IsNull(rsaux3!floa_Sal_Cantidad), 0, rsaux3!floa_Sal_Cantidad)
                                          var_articulo = rsaux3!VCHA_ART_ARTICULO_ID
                                          var_cadena = "INSERT INTO TB_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_AÑO) VALUES "
                                          var_cadena = var_cadena + "('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_surtido + "', '" + var_movimiento_traspaso + "', " + CStr(VAR_FOLIO_TRASPASO) + ", "
                                          var_cadena = var_cadena + "'" + var_articulo + "', " + CStr(var_cantidad) + ",  " + CStr(var_costo) + ", " + CStr(rsaux3!floa_Sal_precio) + ",0, '', 2005)"
                                          rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          rsaux4.Open "INSERT INTO TB_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_SAL_AÑO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_calidad + "', '" + var_movimiento_traspaso + "', " + CStr(VAR_FOLIO_TRASPASO) + ", '" + rsaux3!VCHA_ART_ARTICULO_ID + "', " + CStr(rsaux3!floa_Sal_Cantidad) + ",  " + CStr(var_costo) + ", " + CStr(rsaux3!floa_Sal_precio) + ",0,2005)"
                                          rsaux3.MoveNext
                                    Wend
                                    rsaux3.Close
                                    rsaux3.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET CHAR_EMO_ESTATUS = 'I', DTIM_EMO_FECHA_FINALIZO = GETDATE() WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_calidad + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_traspaso + "' AND INTE_EMO_NUMERO = " + CStr(VAR_FOLIO_TRASPASO), cnn, adOpenDynamic, adLockOptimistic
                                                                       
                                                                       
'''''''''' se genera el pedido y la orden de surtido

                                    rsaux3.Open "SELECT CHAR_TPE_TIPO_PEDIDO_ID, INTE_PLA_DIAS, INTE_TPE_DIAS_CADUCIDAD, FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2, FLOA_GAC_DESCUENTO_3 FROM  VW_PEDIDOS_2 WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_clave_cliente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux3.EOF Then
                                       var_tipo_pedido = IIf(IsNull(rsaux3!char_tpe_tipo_pedido_id), "M", rsaux3!char_tpe_tipo_pedido_id)
                                       var_plazo = IIf(IsNull(rsaux3!inte_pla_dias), 0, rsaux3!inte_pla_dias)
                                       var_caducidad = IIf(IsNull(rsaux3!inte_tpe_dias_caducidad), 0, rsaux3!inte_tpe_dias_caducidad)
                                       var_descuento_1_nuevo = IIf(IsNull(rsaux3!floa_gac_Descuento_1), 0, rsaux3!floa_gac_Descuento_1)
                                       var_descuento_2_nuevo = IIf(IsNull(rsaux3!FLOA_GAC_DESCUENTO_2), 0, rsaux3!FLOA_GAC_DESCUENTO_2)
                                       If Me.txt_clave_cliente_nuevo = "C000005397" Then
                                          var_descuento_1_nuevo = 12
                                          var_descuento_2_nuevo = 3.96
                                       End If

                                       
                                    End If
                                    rsaux3.Close
                                    
                                    rsaux3.Open "SELECT VCHA_TIT_TITULAR_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_AGE_AGENTE_ID, VCHA_LIS_LISTA_ID, VCHA_CAN_CANAL_VENTA_ID, VCHA_MON_MONEDA_ID, ISNULL(INTE_MON_MONEDA_LOCAL,0) as inte_mon_moneda_local, ISNULL(FLOA_GAC_DESCUENTO_1,0) AS FLOA_GAC_DESCUENTO_1, ISNULL(FLOA_GAC_DESCUENTO_2,0) AS FLOA_GAC_DESCUENTO_2 FROM VW_CLIENTES WHERE  VCHA_CLI_CLAVE_ID = '" + Me.txt_clave_cliente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux3.EOF Then
                                       VAR_TITULAR_NUEVO = rsaux3!vcha_tit_titular_id
                                       var_GRUPO_ACTUAL_NUEVO = rsaux3!VCHA_GAC_GRUPO_aCTUAL_ID
                                       var_GRUPO_REAL_NUEVO = rsaux3!vcha_gre_grupo_real_id
                                       VAR_AGENTE_NUEVO = rsaux3!VCHA_AGE_AGENTE_ID
                                       var_lista_precios = rsaux3!vcha_LIS_LISTA_iD
                                       var_canal_venta = rsaux3!vcha_can_canal_venta_id
                                       VAR_MONEDA_NUEVO = rsaux3!vcha_mon_moneda_id
                                       var_moneda_local = rsaux3!inte_mon_moneda_local
                                       var_descuento_1_nuevo = IIf(IsNull(rsaux3!floa_gac_Descuento_1), 0, rsaux3!floa_gac_Descuento_1)
                                       var_descuento_2_nuevo = IIf(IsNull(rsaux3!FLOA_GAC_DESCUENTO_2), 0, rsaux3!FLOA_GAC_DESCUENTO_2)
                                       
                                       If Me.txt_clave_cliente_nuevo = "C000005397" Then
                                          var_descuento_1_nuevo = 12
                                          var_descuento_2_nuevo = 3.96
                                       End If
                                       
                                       
                                       
                                       franquicia = 0
                                       rsaux9.Open "select isnull(inte_esb_franquicia,0) from tb_establecimientos where vcha_Esb_establecimiento_id = '" + Me.txt_clave_establecimiento_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux9.EOF Then
                                          franquicia = IIf(IsNull(rsaux9(0).Value), 0, rsaux9(0).Value)
                                          If franquicia = 1 Then
                                             var_descuento_1_nuevo = 21
                                             var_descuento_1_nuevo = var_descuento_1_nuevo + 2.5
                                          End If
                                       End If
                                       rsaux9.Close
                                    End If
                                    rsaux3.Close

                                    var_TIPO_CAMBIO_NUEVO = 0
                                    If var_moneda_local = 1 Then
                                       var_TIPO_CAMBIO_NUEVO = 1
                                    Else
                                       rsaux3.Open "SELECT * FROM VW_TIPOCAMBIO_FECHA WHERE VCHA_MON_MONEDA_ID = '" + VAR_MONEDA_NUEVO + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                       If Not rsaux3.EOF Then
                                          var_TIPO_CAMBIO_NUEVO = rsaux3!mone_tca_importe
                                       End If
                                       rsaux3.Close
                                    End If
                                   'MsgBox (var_TIPO_CAMBIO_NUEVO)
                                   If var_TIPO_CAMBIO_NUEVO > 0 Then
                                      cnn.BeginTrans
                                      rsaux3.Open "SELECT ISNULL(MAXIMO,0) FROM VW_MAXIMO_PEDIDO", cnn, adOpenDynamic, adLockOptimistic
                                      var_numero_pedido = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value) + 1
                                      rsaux3.Close
                                      
                                      
                                      
                                      rsaux3.Open "SELECT CHAR_TPE_TIPO_PEDIDO_ID, INTE_PLA_DIAS, INTE_TPE_DIAS_CADUCIDAD, FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2, FLOA_GAC_DESCUENTO_3 FROM  VW_PEDIDOS_2 WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_clave_cliente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
                                      If Not rsaux3.EOF Then
                                         var_tipo_pedido = IIf(IsNull(rsaux3!char_tpe_tipo_pedido_id), "M", rsaux3!char_tpe_tipo_pedido_id)
                                         var_plazo = IIf(IsNull(rsaux3!inte_pla_dias), 0, rsaux3!inte_pla_dias)
                                         var_caducidad = IIf(IsNull(rsaux3!inte_tpe_dias_caducidad), 0, rsaux3!inte_tpe_dias_caducidad)
                                         var_descuento_1_nuevo = IIf(IsNull(rsaux3!floa_gac_Descuento_1), 0, rsaux3!floa_gac_Descuento_1)
                                         var_descuento_2_nuevo = IIf(IsNull(rsaux3!FLOA_GAC_DESCUENTO_2), 0, rsaux3!FLOA_GAC_DESCUENTO_2)
                                         If Me.txt_clave_cliente_nuevo = "C000005397" Then
                                            var_descuento_1_nuevo = 12
                                            var_descuento_2_nuevo = 3.96
                                         End If
                                      End If
                                      rsaux3.Close
                                      
                                      
                                       rsaux9.Open "select isnull(inte_esb_franquicia,0) from tb_establecimientos where vcha_Esb_establecimiento_id = '" + Me.txt_clave_establecimiento_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux9.EOF Then
                                          franquicia = IIf(IsNull(rsaux9(0).Value), 0, rsaux9(0).Value)
                                          If franquicia = 1 Then
                                             var_descuento_1_nuevo = 21
                                             var_descuento_1_nuevo = var_descuento_1_nuevo + 2.5
                                          End If
                                       End If
                                       rsaux9.Close
                                      
                                      
                                      
                                      var_cadena = "INSERT INTO TB_ENCABEZADO_PEDIDOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, INTE_PED_REFERENCIA, DTIM_PED_FECHA, DTIM_PED_REFERENCIA, VCHA_AGE_AGENTE_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_PED_RESURTIBLE, INTE_PED_ESPECIALES, CHAR_PED_ESTATUS, FLOA_PED_DESCUENTO_1, FLOA_PED_DESCUENTO_2, FLOA_PED_DESCUENTO_3, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, DTIM_AUD_FECHA, VCHA_MON_MONEDA_ID, INTE_PED_AUTORIZO, VCHA_PED_AUTORIZO, DTIM_PED_AUTORIZO, CHAR_PED_TIPO, INTE_PED_FACTURA_CEROS) VALUES"
                                      var_cadena = var_cadena + " ( '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_surtido + "', '" + var_tipo_pedido + "', " + CStr(var_numero_pedido) + ", 0, GETDATE(), GETDATE(), '" + VAR_AGENTE_NUEVO + "', '" + VAR_TITULAR_NUEVO + "', '" + Me.txt_clave_cliente_nuevo + "', '" + Me.txt_clave_establecimiento_nuevo + "', 0, 0, 'S', " + CStr(var_descuento_1_nuevo) + ", " + CStr(var_descuento_2_nuevo) + ", 0,  " + CStr(var_plazo) + ",  " + CStr(var_caducidad) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDATE(), '" + VAR_MONEDA_NUEVO + " ', 1, '" + var_clave_usuario_global + "', GETDATE(),'R'," + CStr(var_factura_ceros) + ")"
                                      rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                      rsaux3.Open "select max(inte_ors_orden_surtido) from tb_enc_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
                                      var_numero_orden_surtido = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value) + 1
                                      rsaux3.Close
                                      var_cadena = "INSERT INTO TB_ENC_ORDEN_SURTIDO (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, CHAR_ORS_ESTATUS, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, FLOA_ORS_DESCUENTO_3, VCHA_AUD_USAURIO, VCHA_AUD_MAQUINA, DTIM_AUD_FECHA, INTE_ORS_FACTURA_CEROS, VCHA_MON_MONEDA_ID, CHAR_ORS_TIPO)  VALUES"
                                      var_cadena = var_cadena + "('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_tipo_pedido + "', " + CStr(var_numero_pedido) + ", '" + var_almacen_surtido + "', " + CStr(var_numero_orden_surtido) + ", GETDATE(), GETDATE() + " + CStr(var_plazo) + ", '', '" + VAR_TITULAR_NUEVO + "', '" + Me.txt_clave_cliente_nuevo + "', '" + Me.txt_clave_establecimiento_nuevo + "', " + CStr(var_descuento_1_nuevo) + ", " + CStr(var_descuento_2_nuevo) + ", 0, '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDATE(), " + CStr(var_factura_ceros) + ", '" + VAR_MONEDA_NUEVO + "','R')"
                                      rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                      cnn.CommitTrans
                                      
                                      
                                      
                                      var_cadena = "SELECT     SUM(FLOA_SAL_CANTIDAD) AS floa_sal_cantidad, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID,"
                                      var_cadena = var_cadena + " VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO, VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE , DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                      var_cadena = var_cadena + " From dbo.TB_SALIDAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Car_documento = 'FA' and vcha_Ser_Serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + Me.txt_factura + " GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, INTE_CAR_NUMERO, CHAR_SAL_ESTATUS, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, FLOA_SAL_PRECIO_PROMEDIO, VCHA_REE_FOLIO,"
                                      var_cadena = var_cadena + " VCHA_SAL_REFERENCIA, CHAR_PED_TIPO, VCHA_CAT_CATALOGO_ID, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, INTE_SAL_AÑO, INTE_SAL_CONSECUTIVO, INTE_SAL_CONSECUTIVO_FACTURA, VCHA_ART_NUMERO_SERIE, DTIM_INT_FECHA, VCHA_INT_CARGADO, INTE_INT_INTERFACE"
                                      rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                      'rsaux3.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND INTE_CAR_NUMERO = " + CStr(Me.txt_factura), cnn, adOpenDynamic, adLockOptimistic
                                      
                                      
                                      While Not rsaux3.EOF
                                            var_promocion_1 = 0
                                            var_promocion_2 = 0
                                            var_costo = rsaux3!floa_Sal_costo
                                            var_precio_pedido = rsaux3!floa_Sal_precio
                                            'esto se elimino porque las facturas estab tomando el nuevo precio y no el precio con el que habia facturado
                                            'rsaux10.Open "select floa_dli_precio from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_Art_articulo_id = '" + rsaux3!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                            'var_precio_pedido = rsaux10(0).Value * var_TIPO_CAMBIO_NUEVO
                                            'rsaux10.Close
                                            var_promocion_1 = IIf(IsNull(rsaux3!floa_sal_promocion_1), 0, rsaux3!floa_sal_promocion_1)
                                            var_promocion_2 = IIf(IsNull(rsaux3!FLOA_SAL_PROMOCION_2), 0, rsaux3!FLOA_SAL_PROMOCION_2)

                                            var_cadena = "INSERT INTO TB_DETALLE_PEDIDOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_PED_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_PED_PRECIO, FLOA_PED_CANTIDAD, FLOA_PED_CANTIDAD_SURTIDA, FLOA_PED_PROMOCION_1, FLOA_PED_PROMOCION_2) VALUES "
                                            var_cadena = var_cadena + "     ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_surtido + "', " + CStr(var_numero_pedido) + ", '" + rsaux3!VCHA_ART_ARTICULO_ID + "', " + CStr(var_precio_pedido) + ", " + CStr(rsaux3!floa_Sal_Cantidad) + ", 0, " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ")"
                                            rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                            
                                            var_cadena = "INSERT INTO TB_DET_ORDEN_SURTIDO (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, VCHA_ART_ARTICULO_ID, FLOA_ORS_COSTO, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_PEDIDA, FLOA_ORS_EXISTEN, FLOA_ORS_APARTADAS, FLOA_ORS_POSIBLES, FLOA_ORS_CANTIDAD_SURTIR, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_ORS_CANTIDAD_EMPACADA, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, FLOA_ORS_cANTIDAD_SALIDA) VALUES "
                                            var_cadena = var_cadena + " ('" + var_almacen_surtido + "', '" + var_unidad_oranizacional + "', '" + var_almacen_surtido + "', " + CStr(var_numero_orden_surtido) + ", '" + rsaux3!VCHA_ART_ARTICULO_ID + "', " + CStr(var_costo) + ", " + CStr(var_precio_pedido) + ", " + CStr(rsaux3!floa_Sal_Cantidad) + ", " + CStr(rsaux3!floa_Sal_Cantidad) + ", 0, " + CStr(rsaux3!floa_Sal_Cantidad) + ", " + CStr(rsaux3!floa_Sal_Cantidad) + ", " + CStr(rsaux3!floa_Sal_Cantidad) + ", 0, " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ", " + CStr(rsaux3!floa_Sal_Cantidad) + ")"
                                            rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                            
                                            rsaux4.Open "update tb_Existencias set floa_Exi_Cantidad_apartada = floa_exi_cantidad_apartada - " + CStr(rsaux3!floa_Sal_Cantidad) + " where vcha_Art_articulo_id = '" + rsaux3!VCHA_ART_ARTICULO_ID + "' and vcha_alm_almacen_id = '" + var_almacen_surtido + "'", cnn, adOpenDynamic, adLockOptimistic
                                            rsaux3.MoveNext
                                      Wend
                                      rsaux3.Close
                                      
                                      
                                      
                                      
                                      
                                      rsaux3.Open "SELECT INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE FROM VW_FACTURAS_EMBARQUE WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
                                      If Not rsaux3.EOF Then
                                         var_jaula = IIf(IsNull(rsaux3!inte_jau_jaula_id), 0, rsaux3!inte_jau_jaula_id)
                                         var_vehiculo = IIf(IsNull(rsaux3!vcha_veh_vehiculo_id), "", rsaux3!vcha_veh_vehiculo_id)
                                         var_chofer = IIf(IsNull(rsaux3!vcha_cho_chofer_id), "", rsaux3!vcha_cho_chofer_id)
                                         var_cubicaje = IIf(IsNull(rsaux3!floa_emb_cubicaje), "", rsaux3!floa_emb_cubicaje)
                                      Else
                                         var_jaula = 1
                                         var_vehiculo = ""
                                         var_chofer = ""
                                         var_cubicaje = 0
                                      End If
                                      rsaux3.Close
                                      rsaux3.Open "SELECT MAXIMO_EMBARQUE FROM VW_MAXIMO_EMBARQUE WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                                      VAR_EMBARQUE = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value) + 1
                                      rsaux3.Close
                                      var_cadena = "INSERT INTO TB_ENCABEZADO_EMBARQUES (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, INTE_EMB_EMBARQUE, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, VCHA_AGE_AGENTE_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_EMB_TIPO)"
                                      var_cadena = var_cadena + " VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', " + CStr(VAR_EMBARQUE) + ", " + CStr(var_jaula) + ", '" + var_vehiculo + "', '" + VAR_AGENTE_NUEVO + "', GETDATE(), GETDATE(), 'I', '" + var_chofer + "', " + CStr(var_cubicaje) + ", 'R')"
                                      rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   
                                      rsaux3.Open "SELECT INTE_EMO_NUMERO from TB_FOLIOS_MOVIMIENTOS where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "'", cnn, adOpenDynamic, adLockOptimistic
                                      If rsaux3.EOF Then
                                         rsaux4.Open "INSERT INTO TB_FOLIOS_MOVIMIENTOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_movimiento_salida + "',1)", cnn, adOpenDynamic, adLockOptimistic
                                         var_folio_salida = 0
                                      Else
                                         var_folio_salida = IIf(IsNull(rsaux3!INTE_EMO_NUMERO), 0, rsaux3!INTE_EMO_NUMERO)
                                         rsaux4.Open "UPDATE TB_FOLIOS_MOVIMIENTOS SET INTE_EMO_NUMERO = INTE_EMO_NUMERO + 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "'", cnn, adOpenDynamic, adLockOptimistic
                                      End If
                                      rsaux3.Close
                                      var_folio_salida = var_folio_salida + 1
  
                                      rsaux10.Open "select * from tb_agentes where vcha_Can_Canal_venta_id = '" + VAR_AGENTE_NUEVO + "'", cnn, adOpenDynamic, adLockOptimistic
                                      var_canal_venta = ""
                                      If Not rsaux10.EOF Then
                                         var_canal_venta = IIf(IsNull(rsaux10!vcha_can_canal_venta_id), "", rsaux10!vcha_can_canal_venta_id)
                                      Else
                                         var_canal_venta = ""
                                      End If
                                      rsaux10.Close
                                      var_cadena = "INSERT INTO TB_ENCABEZADO_MOVIMIENTOS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, DTIM_EMO_FECHA, INTE_EMO_NUMERO, INTE_EMO_NUMERO_ORIGEN, VCHA_CLI_CLAVE_ID, VCHA_PRO_PROVEEDOR_ID, VCHA_EMO_ALMACEN_ORIGEN, VCHA_EMO_ALMACEN_DESTINO, CHAR_EMO_ESTATUS, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_EMO_FACTURA, VCHA_EMO_MOVIMIENTO_ORIGEN, VCHA_EMO_REFERENCIA, VCHA_ESB_ESTABLECIMIENTO_ID, CHAR_EMO_BLOQUEADO, VCHA_TIT_TITULAR_ID, VCHA_AGE_AGENTE_ID, FLOA_EMO_DESCUENTO_1, FLOA_EMO_DESCUENTO_2, FLOA_EMO_DESCUENTO_3, VCHA_MON_MONEDA_ID, FLOA_EMO_TIPO_CAMBIO, INTE_EMO_FACTURA_CEROS, vcha_can_canal_venta_id) VALUES"
                                      var_cadena = var_cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_surtido + "', '" + var_movimiento_salida + "',  GETDATE(), " + CStr(var_folio_salida) + ", " + CStr(var_numero_orden_surtido) + ", '" + Me.txt_clave_cliente_nuevo + "', '', '" + var_almacen_surtido + "', '', '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '', '', 'REFACTURACION', '" + Me.txt_clave_establecimiento_nuevo + "', '', '" + VAR_TITULAR_NUEVO + "', '" + VAR_AGENTE_NUEVO + "', " + CStr(var_descuento_1_nuevo) + ", " + CStr(var_descuento_2_nuevo) + ", 0, '" + VAR_MONEDA_NUEVO + "', " + CStr(var_TIPO_CAMBIO_NUEVO) + ", " + CStr(var_factura_ceros) + ",'" + var_canal_venta + "')"
                                      rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                      var_cadena = "INSERT INTO TB_DETALLE_EMBARQUES (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_EMB_EMBARQUE, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_EMB_AUTORIZO)  VALUES"
                                      var_cadena = var_cadena + "('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_surtido + "', " + CStr(VAR_EMBARQUE) + ", '" + var_movimiento_salida + "', " + CStr(var_folio_salida) + ", '" + var_clave_usuario_global + "')"
                                      
                                      rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                      
                                      rsaux3.Open "SELECT VCHA_ART_ARTICULO_ID, FLOA_ORS_COSTO, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIDA, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2 FROM TB_DET_ORDEN_SURTIDO WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(var_numero_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                                      While Not rsaux3.EOF
                                            var_precio = IIf(IsNull(rsaux3!floa_ors_precio), 0, rsaux3!floa_ors_precio)
                                            If var_factura_ceros = 1 Then
                                               var_precio = 0
                                            End If
                                            var_promocion_1 = IIf(IsNull(rsaux3!floa_ors_promocion_1), 0, rsaux3!floa_ors_promocion_1)
                                            var_promocion_2 = IIf(IsNull(rsaux3!floa_ors_promocion_2), 0, rsaux3!floa_ors_promocion_2)
                                            'var_precio = var_precio * (1 - (var_promocion_1 / 100))
                                            var_cadena = "INSERT INTO TB_SALIDAS  (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_AÑO, FLOA_SAL_dESCUENTO_1, FLOA_SAL_dESCUENTO_2) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_surtido + "', '" + var_movimiento_salida + "', " + CStr(var_folio_salida) + ", '" + rsaux3!VCHA_ART_ARTICULO_ID + "', " + CStr(rsaux3!FLOA_ORS_CANTIDAD_SURTIDA) + ",  " + CStr(rsaux3!floa_ors_costo) + ", " + CStr(var_precio) + ", 0, " + CStr(var_promocion_1) + ", " + CStr(var_promocion_2) + ", '" + var_tipo_pedido + "', 2005," + CStr(var_descuento_1_nuevo) + "," + CStr(var_descuento_2_nuevo) + ")"
                                            rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                            'MsgBox var_movimiento_salida
                                            rsaux3.MoveNext
                                      Wend
                                      rsaux3.Close
                                      rsaux3.Open "UPDATE TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I' WHERE  VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND  INTE_EMB_EMBARQUE = " + CStr(VAR_EMBARQUE), cnn, adOpenDynamic, adLockOptimistic
                                      rsaux3.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET CHAR_EMO_ESTATUS = 'I', DTIM_EMO_FECHA_FINALIZO = GETDATE(), FLOA_EMO_TIPO_CAMBIO= " + CStr(var_TIPO_CAMBIO_NUEVO) + " WHERE (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID  = '" + var_almacen_surtido + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' AND INTE_EMO_NUMERO = " + CStr(var_folio_salida) + ")"
                                      MsgBox "Se debe facturar el embarque " + CStr(VAR_EMBARQUE), vbOKOnly, "ATENCION"
                                      
                                      Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS_devoluciones.rpt")
                                      reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADAs_devoluciones.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "' and {VW_MOVIMIENTOS_ENTRADAs_devoluciones.VCHA_MOV_MOVIMIENTO_ID} = '" + var_movimiento_calidad + "' AND {VW_MOVIMIENTOS_ENTRADAs_Devoluciones.INTE_EMO_NUMERO} = " + Str(var_numero_folio_calidad) + " and {VW_MOVIMIENTOS_ENTRADAs_Devoluciones.vcha_emp_empresa_id} = '" + var_empresa + "'"
                                      frmvistasprevias.cr.ReportSource = reporte
                                      For ntablas = 1 To reporte.Database.Tables.Count
                                          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                      Next ntablas
                                      frmvistasprevias.cr.ViewReport
                                      frmvistasprevias.Caption = "Reporte de Movimientos"
                                      frmvistasprevias.Show
                                      Set reporte = Nothing
                                      
                                      
                                      
                                      
                                      
                                      
                                      
                                      
                                      
                                   End If
                                                                       
                                                                       
                                                                       
                                                                       
                                                                       
                                 Else
                                    MsgBox "Falta parametrizar la causa de devolución", vbOKOnly, "ATENCION"
                                 End If
                              Else
                                 MsgBox "Falta parametrizar la causa de devolución", vbOKOnly, "ATENCION"
                              End If
                              rsaux.Close
                           Else
                              MsgBox "La tabla de parametros para refacturación no esta actualizada, consulte a sistemas", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "La tabla de parametros para refacturación no esta actualizada, consulte a sistemas", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "La tabla de parametros para refacturación no esta actualizada, consulte a sistemas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "La tabla de parametros para refacturación no esta actualizada, consulte a sistemas", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "La tabla de parametros para refacturación no esta actualizada, consulte a sistemas", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
         End If
         frm_cliente.Visible = False
      Else
         MsgBox "Debe indicar un establecimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Debe de indicar un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_Click()
   frm_cliente.Visible = False
End Sub

Private Sub cmd_imprimir_Click()
Dim si As Integer
   rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
   Me.txt_clave_cliente_nuevo = rs!vcha_cli_clave_id
   Me.txt_nombre_cliente_nuevo = rs!VCHA_CLI_NOMBRE
   txt_domicilio_nuevo = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
   txt_ciudad_nuevo = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
   txt_colonia_nuevo = IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA)
   txt_estado_nuevo = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
   txt_pais_nuevo = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
   txt_cp_nuevo = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
   txt_rfc_nuevo = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
   txt_descuento_1_nuevo = Format(IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1), "###,###,##0.00")
   txt_descuento_2_nuevo = Format(IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2), "###,###,##0.00")
   rs.Close
   frm_cliente.Visible = True
   Me.txt_clave_establecimiento_nuevo = ""
   Me.txt_nombre_establecimiento_nuevo = ""
   txt_clave_cliente_nuevo.Visible = True
   txt_clave_cliente_nuevo.SetFocus
End Sub

Private Sub cmd_nuevo_Click()
   txt_factura = ""
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   txt_domicilio = ""
   txt_colonia = ""
   txt_ciudad = ""
   txt_estado = ""
   txt_pais = ""
   txt_cp = ""
   txt_rfc = ""
   txt_descuento_1 = ""
   txt_descuento_2 = ""
   txt_subimporte = ""
   txt_iva = ""
   txt_importe = ""
   Me.txt_serie = ""
   Me.txt_clave_agente = ""
   Me.txt_nombre_agente = ""
   Me.txt_clave_titular = ""
   Me.txt_nombre_titular = ""
   Me.txt_clave_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
   
   Me.txt_serie.SetFocus
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 500
   Left = 1000
   txt_factura = ""
   txt_clave_cliente = ""
   txt_nombre_cliente = ""
   txt_domicilio = ""
   txt_colonia = ""
   txt_ciudad = ""
   txt_estado = ""
   txt_pais = ""
   txt_cp = ""
   txt_rfc = ""
   txt_descuento_1 = ""
   txt_descuento_2 = ""
   txt_subimporte = ""
   txt_iva = ""
   txt_importe = ""
   frm_cliente.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_cancela_facturas_devolucion)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If var_tipo_lista = 1 Then
            txt_clave_cliente_nuevo = lv_lista.selectedItem
            txt_nombre_cliente_nuevo = lv_lista.selectedItem.SubItems(1)
            txt_clave_cliente_nuevo.SetFocus
         End If
         If var_tipo_lista = 2 Then
            txt_clave_establecimiento_nuevo = lv_lista.selectedItem
            txt_nombre_establecimiento_nuevo = lv_lista.selectedItem.SubItems(1)
            txt_clave_establecimiento_nuevo.SetFocus
         End If
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

Private Sub Text2_Change()

End Sub

Private Sub txt_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_ciudad_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clave_cliente_nuevo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_cliente_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clave_cliente_nuevo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente_nuevo) <> "" Then
      'rs.Open "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + txt_clave_cliente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente_nuevo = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         txt_domicilio_nuevo = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
         txt_ciudad_nuevo = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
         txt_colonia_nuevo = IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA)
         txt_estado_nuevo = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         txt_pais_nuevo = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
         txt_cp_nuevo = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
         txt_rfc_nuevo = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         txt_descuento_1_nuevo = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
         txt_descuento_2_nuevo = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
         
      Else
         MsgBox "Clave de cliente Incorrecta", vbOKOnly, "ATENCION"
         txt_clave_cliente_nuevo = ""
         txt_nombre_cliente_nuevo = ""
         txt_domicilio_nuevo = ""
         txt_ciudad_nuevo = ""
         txt_colonia_nuevo = ""
         txt_estado_nuevo = ""
         txt_pais_nuevo = ""
         txt_cp_nuevo = ""
         txt_rfc_nuevo = ""
         txt_descuento_1_nuevo = 0
         txt_descuento_2_nuevo = 0
         txt_nombre_establecimiento_nuevo = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_clave_establecimiento_nuevo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_establecimiento_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_esb_establecimiento_id,vcha_esb_nombre from vw_establecimientos where vcha_cli_clave_id = '" + txt_clave_cliente_nuevo + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_establecimiento_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clave_establecimiento_nuevo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_cliente_nuevo) <> "" Then
      If Trim(txt_clave_establecimiento_nuevo) <> "" Then
         
         rsaux2.Open "select * from vw_establecimientos where vcha_cli_clave_id = '" + txt_clave_cliente_nuevo + "' and vcha_esb_establecimiento_id = '" + txt_clave_establecimiento_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            txt_nombre_establecimiento_nuevo = rsaux2!VCHA_ESB_NOMBRE
         Else
            MsgBox "Clave de establecimiento incorrecta", vbOKOnly, "ATENCION"
            txt_clave_establecimiento_nuevo = ""
            txt_nombre_establecimiento_nuevo = ""
         End If
         rsaux2.Close
      End If
   Else
      MsgBox "Debe de seleccionar un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_colonia_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_cp_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_cp_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_descuento_1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_descuento_2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_descuento_2_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_domicilio_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_domicilio_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_estado_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_factura_LostFocus()
    If Trim(txt_factura) <> "" Then
       If IsNumeric(txt_factura) Then
          rsaux2.Open "SELECT * FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_cAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + Me.txt_serie + "' AND INTE_CAR_NUMERO = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
          If Not rsaux2.EOF Then
             rs.Open "select * from vw_documentos_impresion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = 'FA' and inte_Car_numero = " + txt_factura + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_Ser_Serie_id = '" + Me.txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                Me.txt_clave_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                If rsaux9.State = 1 Then
                   rsaux9.Close
                End If
                rsaux9.Open "select * from tb_agentes where vcha_age_agente_id = '" + Me.txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux9.EOF Then
                   Me.txt_nombre_agente = IIf(IsNull(rsaux9!VCHA_AGE_NOMBRE), "", rsaux9!VCHA_AGE_NOMBRE)
                Else
                   Me.txt_nombre_agente = ""
                End If
                rsaux9.Close
                Me.txt_clave_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                rsaux9.Open "select  * from tb_titulares where vcha_tit_titular_id = '" + Me.txt_clave_titular + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux9.EOF Then
                   Me.txt_nombre_titular = IIf(IsNull(rsaux9!VCHA_TIT_NOMBRE), "", rsaux9!VCHA_TIT_NOMBRE)
                Else
                   Me.txt_nombre_titular = ""
                End If
                rsaux9.Close
                Me.txt_clave_establecimiento = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id)
                rsaux9.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id = '" + Me.txt_clave_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux9.EOF Then
                   Me.txt_nombre_establecimiento = IIf(IsNull(rsaux9!VCHA_ESB_NOMBRE), "", rsaux9!VCHA_ESB_NOMBRE)
                Else
                   Me.txt_nombre_establecimiento = ""
                End If
                rsaux9.Close
                txt_clave_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                txt_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
                txt_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                txt_colonia = IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA)
                txt_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                txt_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                txt_cp = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                txt_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                var_tipo_Cambio = IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                txt_descuento_1 = Format(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), "###,###,##0.00")
                txt_descuento_2 = Format(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), "###,###,##0.00")
                txt_subimporte = Format(IIf(IsNull(rs!floa_car_subimporte), 0, rs!floa_car_subimporte) / var_tipo_Cambio, "###,###,##0.00")
                txt_iva = Format(IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva) / var_tipo_Cambio, "###,###,##0.00")
                txt_importe = Format(IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / var_tipo_Cambio, "###,###,##0.00")
                var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                txt_clave_cliente_nuevo = txt_clave_cliente
                txt_domicilio_nuevo = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
                txt_ciudad_nuevo = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                txt_colonia_nuevo = IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA)
                txt_estado_nuevo = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                txt_pais_nuevo = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                txt_cp_nuevo = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                txt_rfc_nuevo = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                txt_descuento_1_nuevo = Format(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), "###,###,##0.00")
                txt_descuento_2_nuevo = Format(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), "###,###,##0.00")
             Else
                MsgBox "Número de factura incorrecta", vbOKOnly, "ATENCION"
                var_agente = ""
                txt_factura = ""
                txt_clave_cliente = ""
                txt_nombre_cliente = ""
                txt_domicilio = ""
                txt_ciudad = ""
                txt_estado = ""
                txt_pais = ""
                txt_cp = ""
                txt_rfc = ""
                txt_descuento_1 = ""
                txt_descuento_2 = ""
                txt_subimporte = ""
                txt_iva = ""
                txt_importe = ""
             End If
             rs.Close
          Else
             MsgBox "La factura no existe", vbOKOnly, "ATENCION"
          End If
          rsaux2.Close
       End If
    End If
End Sub

Private Sub txt_nombre_Change()

End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_iva_Change()
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

Private Sub txt_nombre_cliente_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_establecimiento_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_esb_establecimiento_id,vcha_esb_nombre from vw_establecimientos where vcha_cli_clave_id = '" + txt_clave_cliente_nuevo + "' order by vcha_esb_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_establecimiento_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Me.cmd_aceptar.SetFocus
   End If
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_pais_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_rfc_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_rfc_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_subimporte_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub
