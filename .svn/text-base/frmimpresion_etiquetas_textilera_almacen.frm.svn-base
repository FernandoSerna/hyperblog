VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmimpresion_etiquetas_textilera_almacen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de etiquetas desde el AG"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_codigo_vianney 
      Height          =   1050
      Left            =   45
      TabIndex        =   34
      Top             =   525
      Width           =   4425
      Begin VB.TextBox txt_codigo_vianney 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   435
         Left            =   1590
         MaxLength       =   5
         TabIndex        =   35
         Text            =   "88888"
         Top             =   495
         Width           =   960
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   " Código de Vianney"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   45
         TabIndex        =   36
         Top             =   135
         Width           =   4335
      End
   End
   Begin VB.Frame frm_estampados 
      Height          =   2835
      Left            =   45
      TabIndex        =   8
      Top             =   525
      Width           =   4410
      Begin VB.TextBox txt_clave_estampado 
         Height          =   360
         Left            =   75
         TabIndex        =   9
         Top             =   495
         Width           =   4215
      End
      Begin MSComctlLib.ListView lv_estampados 
         Height          =   1830
         Left            =   90
         TabIndex        =   10
         Top             =   915
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   3228
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin VB.Label lbl_estampado 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   4305
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   45
      TabIndex        =   12
      Top             =   1245
      Width           =   4410
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   60
         TabIndex        =   13
         Top             =   390
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   3228
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
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   4305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Código del Artículo "
      Height          =   810
      Left            =   60
      TabIndex        =   32
      Top             =   510
      Width           =   4425
      Begin VB.TextBox txt_tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   90
         MaxLength       =   1
         TabIndex        =   0
         Top             =   255
         Width           =   345
      End
      Begin VB.TextBox txt_division 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   450
         MaxLength       =   2
         TabIndex        =   1
         Top             =   255
         Width           =   585
      End
      Begin VB.TextBox txt_subdivision 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1050
         MaxLength       =   2
         TabIndex        =   2
         Top             =   255
         Width           =   585
      End
      Begin VB.TextBox txt_estampado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1650
         MaxLength       =   5
         TabIndex        =   3
         Top             =   255
         Width           =   1230
      End
      Begin VB.TextBox txt_Descuento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3975
         MaxLength       =   1
         TabIndex        =   4
         Top             =   255
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "descuento"
         Height          =   195
         Left            =   3015
         TabIndex        =   33
         Top             =   375
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   0
      TabIndex        =   31
      Top             =   375
      Width           =   4545
   End
   Begin VB.CommandButton cmd_imprimir 
      Height          =   375
      Left            =   90
      Picture         =   "frmimpresion_etiquetas_textilera_almacen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmd_salir 
      Height          =   375
      Left            =   4110
      Picture         =   "frmimpresion_etiquetas_textilera_almacen.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   " Tipo de etiqueta "
      Height          =   870
      Left            =   75
      TabIndex        =   28
      Top             =   1335
      Visible         =   0   'False
      Width           =   45
      Begin VB.OptionButton opt_sencilla 
         Caption         =   "Etiqueta Sencilla"
         Height          =   270
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1590
      End
      Begin VB.OptionButton opt_bulto 
         Caption         =   "Etiqueta por bulto"
         Height          =   255
         Left            =   225
         TabIndex        =   29
         Top             =   495
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Descripción del Artículo "
      Height          =   1725
      Left            =   45
      TabIndex        =   19
      Top             =   2220
      Width           =   4425
      Begin VB.TextBox txt_tipo_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   23
         Top             =   285
         Width           =   3285
      End
      Begin VB.TextBox txt_division_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   22
         Top             =   630
         Width           =   3285
      End
      Begin VB.TextBox txt_subdivision_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   21
         Top             =   975
         Width           =   3285
      End
      Begin VB.TextBox txt_estampado_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   20
         Top             =   1305
         Width           =   3285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   345
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "División:"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Subdivisión:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estampado:"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   1365
         Width           =   840
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Cantidad "
      Height          =   870
      Left            =   60
      TabIndex        =   18
      Top             =   1335
      Width           =   4425
      Begin VB.TextBox txt_Cantidad 
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
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   5
         Top             =   240
         Width           =   2040
      End
   End
   Begin VB.Frame frm_bulto 
      Height          =   975
      Left            =   1290
      TabIndex        =   15
      Top             =   3705
      Width           =   2160
      Begin VB.TextBox txt_cantidad_bulto 
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
         Left            =   45
         MaxLength       =   5
         TabIndex        =   16
         Top             =   420
         Width           =   2040
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   " Número de Etiquetas"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmimpresion_etiquetas_textilera_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista  As Integer

Private Sub cmd_imprimir_Click()
   Dim var_codigo As String
   Dim VERIFICADOR As Integer
   If Me.opt_sencilla = True Then
      If Trim(txt_tipo) <> "" Then
         If Trim(txt_division) <> "" Then
            If Trim(txt_subdivision) <> "" Then
               If Trim(txt_subdivision) <> "" Then
                  If Trim(txt_estampado) <> "" Then
                     If Trim(Me.txt_Descuento) = "" Then
                        Me.txt_Descuento = "0"
                     End If
                     If Trim(txt_Descuento) <> "" Then
                        var_codigo = Trim(txt_tipo) + Trim(txt_division) + Trim(txt_subdivision) + Trim(txt_estampado) + Trim(txt_Descuento)
                        sum1 = 0
                        sum2 = 0
                        mcodigo = var_codigo
                        longitud = Len(mcodigo)
                        For icont = 1 To longitud
                            If ((icont / 2) - Int((icont / 2))) = 0 Then
                               sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                            Else
                               sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                            End If
                        Next icont
                        msuma = sum1 * 13 + sum2
                        VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                        If VERIFICADOR = 10 Then
                           VERIFICADOR = 0
                        End If
                        var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                        rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + var_codigo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                        If rs.EOF Then
                           VAR_CODIGO_alta = Trim(txt_tipo) + Trim(txt_division) + Trim(txt_subdivision) + Trim(txt_estampado) + "0"
                           sum1 = 0
                           sum2 = 0
                           mcodigo = VAR_CODIGO_alta
                           longitud = Len(mcodigo)
                           For icont = 1 To longitud
                               If ((icont / 2) - Int((icont / 2))) = 0 Then
                                  sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                               Else
                                  sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                               End If
                           Next icont
                           msuma = sum1 * 13 + sum2
                           VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                           If VERIFICADOR = 10 Then
                              VERIFICADOR = 0
                           End If
                           
                           VAR_CODIGO_alta = VAR_CODIGO_alta + Trim(CStr(VERIFICADOR))
                           rsaux.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + VAR_CODIGO_alta + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_precio = IIf(IsNull(rsaux!mone_Art_precio_base), 0, rsaux!mone_Art_precio_base)
                              VAR_dESCUENTO_ARTICULO = (100 - (CDbl(Me.txt_Descuento) * 10)) / 100
                              var_precio = var_precio * VAR_dESCUENTO_ARTICULO
                              var_cadena = "INSERT INTO TB_ARTICULOS (VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_BAJA, VCHA_ART_CATALOGO_INICIO, VCHA_ART_CATALOGO_VIGENTE, VCHA_LIC_LICENCIA_ID, VCHA_ART_NUMERO_LIC, VCHA_DIS_DISEÑO_ID,"
                              var_cadena = var_cadena + " VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID, VCHA_CAR_CLASE_ID, VCHA_ART_ESTAMPADO1, VCHA_ART_TIPO_ESTAMPADO1, VCHA_ART_ESTAMPADO2, VCHA_ART_TIPO_ESTAMPADO2, VCHA_ART_COLOR1, VCHA_ART_COLOR2, VCHA_ART_TONO1, VCHA_ART_TONO2, "
                              var_cadena = var_cadena + " INTE_ART_NUMERO_DECORATIVOS, INTE_ART_FUNDAS, VCHA_USO_USO_ID, VCHA_SUS_SUBTIPO_USO_ID, VCHA_TAL_TALLA_ID, VCHA_UNI_UNIDAD_ID, FLOA_ART_VOLUMEN, FLOA_ART_TELA, VCHA_ART_COMPOSICION, FLOA_ART_PESO, FLOA_ART_TARA, VCHA_CAJ_CAJA_ID, FLOA_ART_PIEZAS_CAJA, FLOA_ART_MAXIMO, "
                              var_cadena = var_cadena + " FLOA_ART_MINIMO, FLOA_ART_PUNTO_REORDEN, FLOA_ART_DIAS_INVENTARIO, VCHA_UBI_UNICACION_ID, FLOA_ART_BULTO, INTE_ART_SALIDA_MASIVA, VCHA_EQU_EQUIVALENCIA_ID, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_ART_DETENIDO, VCHA_ART_CODIGO_EXTERNO, NUM_INTER_TRANC_TYPE, "
                              var_cadena = var_cadena + " NUM_INTER_UPLOADED,DATE_INTER_DATE, VCHA_TPR_TIPO_PRODUCTO_ID, VCHA_DIV_DIVISION_ID, VCHA_SUB_SUBDIVISION_ID, VCHA_EST_ESTAMPADO_ID, FLOA_ART_COSTO_MATERIA_PRIMA_ESTANDAR, FLOA_ART_COSTO_AVIOS_ESTANDAR, FLOA_ART_COSTO_MANO_OBRA_ESTANDAR, FLOA_ART_COSTO_GASTOS_FABRICACION_ESTANDAR,"
                              var_cadena = var_cadena + " FLOA_ART_COSTO_MATERIA_PRIMA_REAL, FLOA_ART_COSTO_AVIOS_REAL, FLOA_ART_COSTO_MANO_OBRA_REAL, FLOA_ART_COSTO_GASTOS_FABRICACION_REAL,  FLOA_ART_COSTO_REAL, FLOA_ART_TIEMPO_FABRICACION, FLOA_ART_VALOR_PRODUCCION, INTE_ART_AGREGA_VALOR_PRODUCCION, VCHA_PRO_PROVEEDOR_ID, FLOA_ART_ULTIMO_COSTO, "
                              var_cadena = var_cadena + " INTE_ART_TIPO_MATERIA_PRIMA) SELECT '" + var_codigo + "', VCHA_ART_NOMBRE_ESPAÑOL, " + CStr(var_precio) + ", MONE_ART_COSTO_ESTANDAR, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_BAJA, VCHA_ART_CATALOGO_INICIO, VCHA_ART_CATALOGO_VIGENTE, VCHA_LIC_LICENCIA_ID, VCHA_ART_NUMERO_LIC, VCHA_DIS_DISEÑO_ID,"
                              var_cadena = var_cadena + " VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID, VCHA_CAR_CLASE_ID, VCHA_ART_ESTAMPADO1, VCHA_ART_TIPO_ESTAMPADO1, VCHA_ART_ESTAMPADO2, VCHA_ART_TIPO_ESTAMPADO2, VCHA_ART_COLOR1, VCHA_ART_COLOR2, VCHA_ART_TONO1, VCHA_ART_TONO2, "
                              var_cadena = var_cadena + " INTE_ART_NUMERO_DECORATIVOS, INTE_ART_FUNDAS, VCHA_USO_USO_ID, VCHA_SUS_SUBTIPO_USO_ID, VCHA_TAL_TALLA_ID, VCHA_UNI_UNIDAD_ID, FLOA_ART_VOLUMEN, FLOA_ART_TELA, VCHA_ART_COMPOSICION, FLOA_ART_PESO, FLOA_ART_TARA, VCHA_CAJ_CAJA_ID, FLOA_ART_PIEZAS_CAJA, FLOA_ART_MAXIMO, "
                              var_cadena = var_cadena + " FLOA_ART_MINIMO, FLOA_ART_PUNTO_REORDEN, FLOA_ART_DIAS_INVENTARIO, VCHA_UBI_UNICACION_ID, FLOA_ART_BULTO, INTE_ART_SALIDA_MASIVA, VCHA_EQU_EQUIVALENCIA_ID, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_ART_DETENIDO, VCHA_ART_CODIGO_EXTERNO, NUM_INTER_TRANC_TYPE, "
                              var_cadena = var_cadena + " NUM_INTER_UPLOADED,DATE_INTER_DATE, VCHA_TPR_TIPO_PRODUCTO_ID, VCHA_DIV_DIVISION_ID, VCHA_SUB_SUBDIVISION_ID, VCHA_EST_ESTAMPADO_ID, FLOA_ART_COSTO_MATERIA_PRIMA_ESTANDAR, FLOA_ART_COSTO_AVIOS_ESTANDAR, FLOA_ART_COSTO_MANO_OBRA_ESTANDAR, FLOA_ART_COSTO_GASTOS_FABRICACION_ESTANDAR,"
                              var_cadena = var_cadena + " FLOA_ART_COSTO_MATERIA_PRIMA_REAL, FLOA_ART_COSTO_AVIOS_REAL, FLOA_ART_COSTO_MANO_OBRA_REAL, FLOA_ART_COSTO_GASTOS_FABRICACION_REAL,  FLOA_ART_COSTO_REAL, FLOA_ART_TIEMPO_FABRICACION, FLOA_ART_VALOR_PRODUCCION, INTE_ART_AGREGA_VALOR_PRODUCCION, VCHA_PRO_PROVEEDOR_ID, FLOA_ART_ULTIMO_COSTO, "
                              var_cadena = var_cadena + " INTE_ART_TIPO_MATERIA_PRIMA from tb_articulos WHERE VCHA_ART_ARTICULO_ID = '" + VAR_CODIGO_alta + "'"
                              rsaux2.Open var_cadena, cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux.Close
                        End If
                        rs.Close
                        rsaux5.Open "select * from tb_articulos where vcha_art_Articulo_id = '" + VAR_CODIGO_alta + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                           rsaux.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id = '" + var_codigo + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                           If rsaux.EOF Then
                              rsaux2.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id = '" + VAR_CODIGO_alta + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 var_precio = IIf(IsNull(rsaux2!floa_dli_precio), 0, rsaux2!floa_dli_precio)
                                 VAR_dESCUENTO_ARTICULO = (100 - (CDbl(Me.txt_Descuento) * 10)) / 100
                                 var_precio = var_precio * VAR_dESCUENTO_ARTICULO
                              End If
                              rsaux2.Close
                              rsaux2.Open "insert into tb_detalle_lista_precios (vcha_art_articulo_id, vcha_lis_lista_precios_id, floa_dli_precio) values ('" + var_codigo + "','01'," + CStr(var_precio) + ")", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux.Close
                        End If
                        rsaux5.Close
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + var_codigo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           If CInt(Me.txt_Cantidad) > 0 Then
                              t11 = Trim(Me.txt_subdivision_descripcion)
                              t22 = Trim(Me.txt_estampado_descripcion)
                              t33 = Trim(Me.txt_Descuento)
                              s = Trim(var_codigo)
                              If CInt(Me.txt_Descuento) > 0 Then
                                 Open (App.Path & "\etiqueta.bat") For Output As #2
                                 Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                                 Open (App.Path & "\etiqueta.txt") For Output As #1
                                 Close #2
                                 For i = 1 To CInt(Me.txt_Cantidad)
                                     Print #1, "US"
                                     Print #1, "N"
                                     Print #1, "Q256,24"
                                     Print #1, "q512"
                                     Print #1, "A80,20,0,3,1,1,N,""" + t11 + """"
                                     Print #1, "A80,50,0,3,1,1,N,""" + t22 + """"
                                     Print #1, "A80,80,0,3,1,1,N,""" + "Descuento:" + """"
                                     Print #1, "A300,63,0,5,1,1,N,""" + t33 + "0%"""
                                     Print #1, "B80,120,0,3,2,4,80,B,""" + s + """"
                                     Print #1, "P1"
                                 Next i
                                 'Print #1, ""
                                 Close #1
                                 x = Shell(App.Path & "\etiqueta.bat", vbHide)
                                 'Shell ("print /d:LPT1 " + App.Path + "\etiqueta.txt")
                                 'Shell ("copy " + App.Path + "\etiqueta.txt lpt1 ")
                              Else
                                 Open (App.Path & "\etiqueta.bat") For Output As #2
                                 Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                                 Open (App.Path & "\etiqueta.txt") For Output As #1
                                 Close #2
                                 For i = 1 To CInt(Me.txt_Cantidad)
                                     Print #1, "US"
                                     Print #1, "N"
                                     Print #1, "Q256,24"
                                     Print #1, "q512"
                                     Print #1, "A80,20,0,3,1,1,N,""" + t11 + """"
                                     Print #1, "A80,50,0,3,1,1,N,""" + t22 + """"
                                     Print #1, "B80,120,0,3,2,4,80,B,""" + s + """"
                                     Print #1, "P1"
                                 Next i
                                 'Print #1, ""
                                 Close #1
                                 z = "copy " + App.Path & "\etiqueta.txt lpt1"
                                 x = Shell(App.Path & "\etiqueta.bat", vbHide)
                                 'Shell ("copy " + App.Path + "\etiqueta.txt lpt1 ")
                                 'Shell ("print /d:LPT1 " + App.Path + "\etiqueta.txt")
                              End If
                              Me.txt_tipo = ""
                              Me.txt_tipo_descripcion = ""
                              Me.txt_division = ""
                              Me.txt_division_descripcion = ""
                              Me.txt_subdivision = ""
                              Me.txt_subdivision_descripcion = ""
                              Me.txt_estampado = ""
                              Me.txt_estampado_descripcion = ""
                              Me.txt_Descuento = ""
                              Me.txt_Cantidad = ""
                              Me.txt_tipo.SetFocus
                              
                           End If
                        Else
                        
                           var_equivalencia = Me.txt_estampado
                           If rsaux9.State = 1 Then
                              rsaux9.Close
                           End If
                           rsaux9.Open "select * from tb_Articulos where substring(vcha_art_articulo_id,7,5) = '" + Trim(var_equivalencia) + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux9.EOF Then
                              var_linea = IIf(IsNull(rsaux9!vcha_lin_linea_id), "", rsaux9!vcha_lin_linea_id)
                              var_costo = IIf(IsNull(rsaux9!mone_Art_costo_estandar), 0, rsaux9!mone_Art_costo_estandar)
                              var_precio = IIf(IsNull(rsaux9!mone_Art_precio_base), 0, rsaux9!mone_Art_precio_base)
                              var_catalogo = IIf(IsNull(rsaux9!vcha_Art_catalogo_vigente), "", rsaux9!vcha_Art_catalogo_vigente)
                              var_DEscripcion = IIf(IsNull(rsaux9!vcha_Art_nombre_español), "", rsaux9!vcha_Art_nombre_español)
                              var_proveedor = ""
                              var_linea = Me.txt_division
                              If var_linea = "00" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "2" Then
                                 linea_textilera = "30"
                              End If
                              If var_linea = "10" Then
                                 linea_textilera = "12"
                              End If
                              If var_linea = "11" Then
                                 linea_textilera = "75"
                              End If
                              If var_linea = "12" Then
                                 linea_textilera = "10"
                              End If
                              If var_linea = "13" Then
                                 linea_textilera = "30"
                              End If
                              If var_linea = "14" Then
                                 linea_textilera = "40"
                              End If
                              If var_linea = "15" Then
                                 linea_textilera = "50"
                              End If
                              If var_linea = "16" Then
                                 linea_textilera = "20"
                              End If
                              If var_linea = "20" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "22" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "23" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "24" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "28" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "29" Then
                                 linea_textilera = "12"
                              End If
                              If var_linea = "30" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "31" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "35" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "39" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "40" Then
                                 linea_textilera = "14"
                              End If
                              If var_linea = "41" Then
                                 linea_textilera = "14"
                              End If
                              If var_linea = "42" Then
                                 linea_textilera = "15"
                              End If
                              If var_linea = "43" Then
                                 linea_textilera = "15"
                              End If
                              If var_linea = "44" Then
                                 linea_textilera = "25"
                              End If
                              If var_linea = "45" Then
                                 linea_textilera = "24"
                              End If
                              If var_linea = "50" Then
                                 linea_textilera = "15"
                              End If
                              If var_linea = "55" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "59" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "60" Then
                                 linea_textilera = "14"
                              End If
                              If var_linea = "65" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "70" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "75" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "80" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "90" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "91" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "92" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "93" Then
                                 linea_textilera = "16"
                              End If
                              If var_linea = "94" Then
                                 linea_textilera = "13"
                              End If
                              If var_linea = "95" Then
                                 linea_textilera = "13"
                              End If
                              var_linea = Me.txt_division
                              var_codigo = var_equivalencia
                              var_codigo_textilera = "6" + var_linea + Me.txt_subdivision + var_codigo + "0"
                              var_codigo = var_codigo_textilera
                              sum1 = 0
                              sum2 = 0
                              mcodigo = var_codigo
                              longitud = Len(mcodigo)
                              For icont = 1 To longitud
                                  If ((icont / 2) - Int((icont / 2))) = 0 Then
                                     sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                                  Else
                                     sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                                  End If
                              Next icont
                              msuma = sum1 * 13 + sum2
                              VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                              If VERIFICADOR = 10 Then
                                 VERIFICADOR = 0
                              End If
                              var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                              var_codigo_textilera = var_codigo
                              If rs.State = 1 Then
                                 rs.Close
                              End If
                              rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_aRTICULO_ID = '" + var_codigo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                              If rs.EOF Then
                                 var_cadena = "INSERT INTO TB_ARTICULOS (VCHA_aRT_aRTICULO_ID, VCHA_aRT_nombre_español, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR,        VCHA_LIN_LINEA_ID, VCHA_ART_CATALOGO_VIGENTE,       VCHA_TPR_TIPO_PRODUCTO_ID,               VCHA_DIV_DIVISION_ID,                     VCHA_SUB_SUBDIVISION_ID,                     VCHA_EST_ESTAMPADO_ID,         VCHA_eMP_eMPRESA_ID) VALUES"
                                 var_cadena = var_cadena + "('" + var_codigo_textilera + "', '" + var_DEscripcion + "', " + CStr(var_precio) + ", " + CStr(var_costo) + ",    '" + var_linea + "',  '" + var_catalogo + "', '" + Mid(var_codigo_textilera, 1, 1) + "','" + Mid(var_codigo_textilera, 2, 2) + "', '" + Mid(var_codigo_textilera, 4, 2) + "', '" + Mid(var_codigo_textilera, 6, 5) + "','18')"
                                 rsaux.Open var_cadena, cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 rsaux.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('01','" + var_codigo_textilera + "', " + CStr(var_precio) + ")", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 rsaux.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Mid(var_codigo_textilera, 6, 5) + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 If rsaux.EOF Then
                                    rsaux2.Open "insert into tb_estampados (vcha_est_estampado_id, vcha_est_nombre) values ('" + Mid(var_codigo_textilera, 6, 5) + "', '" + var_DEscripcion + "')", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux.Close
                                 rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 If rsaux.EOF Then
                                    rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id) values ('" + var_equivalencia + "', '" + var_codigo_textilera + "')", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux.Close
                              End If
                              rs.Close
                              
''''''' impresion

                              var_codigo = Trim(txt_tipo) + Trim(txt_division) + Trim(txt_subdivision) + Trim(txt_estampado) + Trim(txt_Descuento)
                              sum1 = 0
                              sum2 = 0
                              mcodigo = var_codigo
                              longitud = Len(mcodigo)
                              For icont = 1 To longitud
                                  If ((icont / 2) - Int((icont / 2))) = 0 Then
                                     sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                                  Else
                                     sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                                  End If
                              Next icont
                              msuma = sum1 * 13 + sum2
                              VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                              If VERIFICADOR = 10 Then
                                 VERIFICADOR = 0
                              End If
                              var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                              rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + var_codigo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                              If rs.EOF Then
                                 VAR_CODIGO_alta = Trim(txt_tipo) + Trim(txt_division) + Trim(txt_subdivision) + Trim(txt_estampado) + "0"
                                 sum1 = 0
                                 sum2 = 0
                                 mcodigo = VAR_CODIGO_alta
                                 longitud = Len(mcodigo)
                                 For icont = 1 To longitud
                                     If ((icont / 2) - Int((icont / 2))) = 0 Then
                                        sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                                     Else
                                        sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                                     End If
                                 Next icont
                                 msuma = sum1 * 13 + sum2
                                 VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                                 If VERIFICADOR = 10 Then
                                    VERIFICADOR = 0
                                 End If
                           
                                 VAR_CODIGO_alta = VAR_CODIGO_alta + Trim(CStr(VERIFICADOR))
                                 rsaux.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + VAR_CODIGO_alta + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 If Not rsaux.EOF Then
                                    var_precio = IIf(IsNull(rsaux!mone_Art_precio_base), 0, rsaux!mone_Art_precio_base)
                                    VAR_dESCUENTO_ARTICULO = (100 - (CDbl(Me.txt_Descuento) * 10)) / 100
                                    var_precio = var_precio * VAR_dESCUENTO_ARTICULO
                                    var_cadena = "INSERT INTO TB_ARTICULOS (VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_BAJA, VCHA_ART_CATALOGO_INICIO, VCHA_ART_CATALOGO_VIGENTE, VCHA_LIC_LICENCIA_ID, VCHA_ART_NUMERO_LIC, VCHA_DIS_DISEÑO_ID,"
                                    var_cadena = var_cadena + " VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID, VCHA_CAR_CLASE_ID, VCHA_ART_ESTAMPADO1, VCHA_ART_TIPO_ESTAMPADO1, VCHA_ART_ESTAMPADO2, VCHA_ART_TIPO_ESTAMPADO2, VCHA_ART_COLOR1, VCHA_ART_COLOR2, VCHA_ART_TONO1, VCHA_ART_TONO2, "
                                    var_cadena = var_cadena + " INTE_ART_NUMERO_DECORATIVOS, INTE_ART_FUNDAS, VCHA_USO_USO_ID, VCHA_SUS_SUBTIPO_USO_ID, VCHA_TAL_TALLA_ID, VCHA_UNI_UNIDAD_ID, FLOA_ART_VOLUMEN, FLOA_ART_TELA, VCHA_ART_COMPOSICION, FLOA_ART_PESO, FLOA_ART_TARA, VCHA_CAJ_CAJA_ID, FLOA_ART_PIEZAS_CAJA, FLOA_ART_MAXIMO, "
                                    var_cadena = var_cadena + " FLOA_ART_MINIMO, FLOA_ART_PUNTO_REORDEN, FLOA_ART_DIAS_INVENTARIO, VCHA_UBI_UNICACION_ID, FLOA_ART_BULTO, INTE_ART_SALIDA_MASIVA, VCHA_EQU_EQUIVALENCIA_ID, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_ART_DETENIDO, VCHA_ART_CODIGO_EXTERNO, NUM_INTER_TRANC_TYPE, "
                                    var_cadena = var_cadena + " NUM_INTER_UPLOADED,DATE_INTER_DATE, VCHA_TPR_TIPO_PRODUCTO_ID, VCHA_DIV_DIVISION_ID, VCHA_SUB_SUBDIVISION_ID, VCHA_EST_ESTAMPADO_ID, FLOA_ART_COSTO_MATERIA_PRIMA_ESTANDAR, FLOA_ART_COSTO_AVIOS_ESTANDAR, FLOA_ART_COSTO_MANO_OBRA_ESTANDAR, FLOA_ART_COSTO_GASTOS_FABRICACION_ESTANDAR,"
                                    var_cadena = var_cadena + " FLOA_ART_COSTO_MATERIA_PRIMA_REAL, FLOA_ART_COSTO_AVIOS_REAL, FLOA_ART_COSTO_MANO_OBRA_REAL, FLOA_ART_COSTO_GASTOS_FABRICACION_REAL,  FLOA_ART_COSTO_REAL, FLOA_ART_TIEMPO_FABRICACION, FLOA_ART_VALOR_PRODUCCION, INTE_ART_AGREGA_VALOR_PRODUCCION, VCHA_PRO_PROVEEDOR_ID, FLOA_ART_ULTIMO_COSTO, "
                                    var_cadena = var_cadena + " INTE_ART_TIPO_MATERIA_PRIMA) SELECT '" + var_codigo + "', VCHA_ART_NOMBRE_ESPAÑOL, " + CStr(var_precio) + ", MONE_ART_COSTO_ESTANDAR, DTIM_ART_FECHA_ALTA, DTIM_ART_FECHA_BAJA, VCHA_ART_CATALOGO_INICIO, VCHA_ART_CATALOGO_VIGENTE, VCHA_LIC_LICENCIA_ID, VCHA_ART_NUMERO_LIC, VCHA_DIS_DISEÑO_ID,"
                                    var_cadena = var_cadena + " VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID, VCHA_CAR_CLASE_ID, VCHA_ART_ESTAMPADO1, VCHA_ART_TIPO_ESTAMPADO1, VCHA_ART_ESTAMPADO2, VCHA_ART_TIPO_ESTAMPADO2, VCHA_ART_COLOR1, VCHA_ART_COLOR2, VCHA_ART_TONO1, VCHA_ART_TONO2, "
                                    var_cadena = var_cadena + " INTE_ART_NUMERO_DECORATIVOS, INTE_ART_FUNDAS, VCHA_USO_USO_ID, VCHA_SUS_SUBTIPO_USO_ID, VCHA_TAL_TALLA_ID, VCHA_UNI_UNIDAD_ID, FLOA_ART_VOLUMEN, FLOA_ART_TELA, VCHA_ART_COMPOSICION, FLOA_ART_PESO, FLOA_ART_TARA, VCHA_CAJ_CAJA_ID, FLOA_ART_PIEZAS_CAJA, FLOA_ART_MAXIMO, "
                                    var_cadena = var_cadena + " FLOA_ART_MINIMO, FLOA_ART_PUNTO_REORDEN, FLOA_ART_DIAS_INVENTARIO, VCHA_UBI_UNICACION_ID, FLOA_ART_BULTO, INTE_ART_SALIDA_MASIVA, VCHA_EQU_EQUIVALENCIA_ID, '18', VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, INTE_ART_DETENIDO, VCHA_ART_CODIGO_EXTERNO, NUM_INTER_TRANC_TYPE, "
                                    var_cadena = var_cadena + " NUM_INTER_UPLOADED,DATE_INTER_DATE, VCHA_TPR_TIPO_PRODUCTO_ID, VCHA_DIV_DIVISION_ID, VCHA_SUB_SUBDIVISION_ID, VCHA_EST_ESTAMPADO_ID, FLOA_ART_COSTO_MATERIA_PRIMA_ESTANDAR, FLOA_ART_COSTO_AVIOS_ESTANDAR, FLOA_ART_COSTO_MANO_OBRA_ESTANDAR, FLOA_ART_COSTO_GASTOS_FABRICACION_ESTANDAR,"
                                    var_cadena = var_cadena + " FLOA_ART_COSTO_MATERIA_PRIMA_REAL, FLOA_ART_COSTO_AVIOS_REAL, FLOA_ART_COSTO_MANO_OBRA_REAL, FLOA_ART_COSTO_GASTOS_FABRICACION_REAL,  FLOA_ART_COSTO_REAL, FLOA_ART_TIEMPO_FABRICACION, FLOA_ART_VALOR_PRODUCCION, INTE_ART_AGREGA_VALOR_PRODUCCION, VCHA_PRO_PROVEEDOR_ID, FLOA_ART_ULTIMO_COSTO, "
                                    var_cadena = var_cadena + " INTE_ART_TIPO_MATERIA_PRIMA from tb_articulos WHERE VCHA_ART_ARTICULO_ID = '" + VAR_CODIGO_alta + "'"
                                    rsaux2.Open var_cadena, cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux.Close
                              End If
                              rs.Close
                              rsaux5.Open "select * from tb_articulos where vcha_art_Articulo_id = '" + VAR_CODIGO_alta + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                              If Not rsaux5.EOF Then
                                 rsaux.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id = '" + var_codigo + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 If rsaux.EOF Then
                                    rsaux2.Open "select * from tb_detalle_lista_precios where vcha_art_articulo_id = '" + VAR_CODIGO_alta + "' and vcha_lis_lista_precios_id = '01'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                    If Not rsaux2.EOF Then
                                       var_precio = IIf(IsNull(rsaux2!floa_dli_precio), 0, rsaux2!floa_dli_precio)
                                       VAR_dESCUENTO_ARTICULO = (100 - (CDbl(Me.txt_Descuento) * 10)) / 100
                                       var_precio = var_precio * VAR_dESCUENTO_ARTICULO
                                    End If
                                    rsaux2.Close
                                    rsaux2.Open "insert into tb_detalle_lista_precios (vcha_art_articulo_id, vcha_lis_lista_precios_id, floa_dli_precio) values ('" + var_codigo + "','01'," + CStr(var_precio) + ")", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux.Close
                              End If
                              rsaux5.Close
                              If rs.State = 1 Then
                                 rs.Close
                              End If
                              rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + var_codigo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                              If Not rs.EOF Then
                                 If CInt(Me.txt_Cantidad) > 0 Then
                                    t11 = Trim(Me.txt_subdivision_descripcion)
                                    t22 = Trim(Me.txt_estampado_descripcion)
                                    t33 = Trim(Me.txt_Descuento)
                                    s = Trim(var_codigo)
                                    If CInt(Me.txt_Descuento) > 0 Then
                                       Open (App.Path & "\etiqueta.bat") For Output As #2
                                       Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                                       Open (App.Path & "\etiqueta.txt") For Output As #1
                                       Close #2
                                       For i = 1 To CInt(Me.txt_Cantidad)
                                           Print #1, "US"
                                           Print #1, "N"
                                           Print #1, "Q256,24"
                                           Print #1, "q512"
                                           Print #1, "A80,20,0,3,1,1,N,""" + t11 + """"
                                           Print #1, "A80,50,0,3,1,1,N,""" + t22 + """"
                                           Print #1, "A80,80,0,3,1,1,N,""" + "Descuento:" + """"
                                           Print #1, "A300,63,0,5,1,1,N,""" + t33 + "0%"""
                                           Print #1, "B80,120,0,3,2,4,80,B,""" + s + """"
                                           Print #1, "P1"
                                       Next i
                                       'Print #1, ""
                                       Close #1
                                       x = Shell(App.Path & "\etiqueta.bat", vbHide)
                                       x = x
                                       'Shell ("print /d:LPT1 " + App.Path + "\etiqueta.txt")
                                       'Shell ("copy " + App.Path + "\etiqueta.txt lpt1 ")
                                    Else
                                       Open (App.Path & "\etiqueta.bat") For Output As #2
                                       Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
                                       Open (App.Path & "\etiqueta.txt") For Output As #1
                                       Close #2
                                       For i = 1 To CInt(Me.txt_Cantidad)
                                           Print #1, "US"
                                           Print #1, "N"
                                           Print #1, "Q256,24"
                                           Print #1, "q512"
                                           Print #1, "A80,20,0,3,1,1,N,""" + t11 + """"
                                           Print #1, "A80,50,0,3,1,1,N,""" + t22 + """"
                                           Print #1, "B80,120,0,3,2,4,80,B,""" + s + """"
                                           Print #1, "P1"
                                       Next i
                                       'Print #1, ""
                                       Close #1
                                       z = "copy " + App.Path & "\etiqueta.txt lpt1"
                                       x = Shell(App.Path & "\etiqueta.bat", vbHide)
                                       'Shell ("copy " + App.Path + "\etiqueta.txt lpt1 ")
                                       'Shell ("print /d:LPT1 " + App.Path + "\etiqueta.txt")
                                    End If
                                    Me.txt_tipo = ""
                                    Me.txt_tipo_descripcion = ""
                                    Me.txt_division = ""
                                    Me.txt_division_descripcion = ""
                                    Me.txt_subdivision = ""
                                    Me.txt_subdivision_descripcion = ""
                                    Me.txt_estampado = ""
                                    Me.txt_estampado_descripcion = ""
                                    Me.txt_Descuento = ""
                                    Me.txt_Cantidad = ""
                                    Me.txt_tipo.SetFocus
                                
                                 End If
                              End If




'''''''termina impresion
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              Me.txt_tipo = ""
                              Me.txt_tipo_descripcion = ""
                              Me.txt_division = ""
                              Me.txt_division_descripcion = ""
                              Me.txt_subdivision = ""
                              Me.txt_subdivision_descripcion = ""
                              Me.txt_estampado = ""
                              Me.txt_estampado_descripcion = ""
                              Me.txt_Descuento = ""
                              Me.txt_Cantidad = ""
                              Me.txt_tipo.SetFocus
                           
                           
                           
                           Else
                              MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                              Me.txt_tipo = ""
                              Me.txt_tipo_descripcion = ""
                              Me.txt_division = ""
                              Me.txt_division_descripcion = ""
                              Me.txt_subdivision = ""
                              Me.txt_subdivision_descripcion = ""
                              Me.txt_estampado = ""
                              Me.txt_estampado_descripcion = ""
                              Me.txt_Descuento = ""
                              Me.txt_Cantidad = ""
                              Me.txt_tipo.SetFocus
                           End If
                           rsaux9.Close
                        End If
                        rs.Close
                     End If
                  Else
                     MsgBox "No se a seleccionado un estampado", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se a seleccionado una subdivisión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a seleccionado una subdivisión", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado una división", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un tipo de producto", vbOKOnly, "ATENCION"
      End If
   End If
   If opt_bulto = True Then
      Me.frm_bulto.Visible = True
      Me.txt_cantidad_bulto = ""
      Me.txt_cantidad_bulto.SetFocus
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1700
   Left = 3200
   opt_sencilla.Value = True
   Me.frm_bulto.Visible = False
   Me.frm_lista.Visible = False
   Me.frm_estampados.Visible = False
   Me.frm_codigo_vianney.Visible = False
   If var_clave_usuario_global = "U0000000108" Then
      Me.txt_Descuento.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_estampados_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_estampados.ListItems.Count > 0 Then
         Me.txt_estampado = lv_estampados.selectedItem
         Me.txt_estampado.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_estampados.Visible = False
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      If lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 1 Then
            Me.txt_tipo = lv_lista.selectedItem
            Me.txt_tipo.SetFocus
         End If
         If var_tipo_lista = 2 Then
            Me.txt_division = lv_lista.selectedItem
            Me.txt_division.SetFocus
         End If
         If var_tipo_lista = 3 Then
            Me.txt_subdivision = lv_lista.selectedItem
            Me.txt_subdivision.SetFocus
         End If
         If var_tipo_lista = 4 Then
            Me.txt_estampado = lv_lista.selectedItem
            Me.txt_estampado.SetFocus
         End If
      Else
         If var_tipo_lista = 1 Then
            Me.txt_tipo.SetFocus
         End If
         If var_tipo_lista = 2 Then
            Me.txt_division.SetFocus
         End If
         If var_tipo_lista = 3 Then
            Me.txt_subdivision.SetFocus
         End If
         If var_tipo_lista = 4 Then
            Me.txt_estampado.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         Me.txt_tipo.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_division.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_subdivision.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_estampado.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_cantidad_bulto_LostFocus()
   Me.frm_bulto.Visible = False
   
   rs.Open "select * from tb_equivalencias where vcha_art "
End Sub

Private Sub txt_Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_Descuento.SetFocus
   End If
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_clave_estampado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT * from TB_ESTAMPADOS where vcha_est_nombre like '%" + Trim(Me.txt_clave_estampado) + "%' ORDER BY VCHA_EST_nombre", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      Me.lv_estampados.ListItems.Clear
      While Not rs.EOF
            Set list_item = lv_estampados.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_estampado = "ESTAMPADOS"
      var_tipo_lista = 4
      Dim var_n As Integer
      var_n = lv_estampados.ListItems.Count
      If var_n > 6 Then
         lv_estampados.ColumnHeaders(1).Width = 0
         lv_estampados.ColumnHeaders(2).Width = 3800
      Else
         lv_estampados.ColumnHeaders(1).Width = 0
         lv_estampados.ColumnHeaders(2).Width = 3800
      End If
      lv_estampados.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_estampados.Visible = False
   End If
End Sub

Private Sub txt_codigo_vianney_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rsaux8.Open "select * from tb_Articulos where vcha_est_estampado_id = '" + Me.txt_codigo_vianney + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      If Not rsaux8.EOF Then
         Me.txt_tipo = Mid(rsaux8!vcha_Art_Articulo_id, 1, 1)
         Me.txt_division = Mid(rsaux8!vcha_Art_Articulo_id, 2, 2)
         Me.txt_subdivision = Mid(rsaux8!vcha_Art_Articulo_id, 4, 2)
         Me.txt_estampado = Mid(rsaux8!vcha_Art_Articulo_id, 6, 5)
         rsaux7.Open "select * from tb_tipos_productos where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
         If Not rsaux7.EOF Then
            Me.txt_tipo_descripcion = IIf(IsNull(rsaux7!vcha_tpr_nombre), "", rsaux7!vcha_tpr_nombre)
            rsaux7.Close
            rsaux7.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_div_DIVISION_ID = '" + Me.txt_division + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
            If Not rsaux7.EOF Then
               Me.txt_division_descripcion = IIf(IsNull(rsaux7!vcha_div_nombre), "", rsaux7!vcha_div_nombre)
               rsaux7.Close
               rsaux7.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' AND VCHA_SUB_SUBDIVISION_ID = '" + Me.txt_subdivision + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  Me.txt_subdivision_descripcion = IIf(IsNull(rsaux7!vcha_sub_nombre), "", rsaux7!vcha_sub_nombre)
                  rsaux7.Close
                  rsaux7.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_eST_eSTAMPADO_ID = '" + Me.txt_estampado + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     Me.txt_estampado_descripcion = IIf(IsNull(rsaux7!vcha_est_nombre), "", rsaux7!vcha_est_nombre)
                     rsaux7.Close
                  Else
                     rsaux7.Close
                     MsgBox "El estampado no existe", vbOKOnly, "ATENCION"
                     Me.txt_estampado = ""
                     Me.txt_estampado_descripcion = ""
                  End If
               Else
                  rsaux7.Close
                  MsgBox "La subdivisión no existe", vbOKOnly, "ATENCION"
                  Me.txt_subdivision = ""
                  Me.txt_subdivision_descripcion = ""
                  Me.txt_estampado = ""
                  Me.txt_estampado_descripcion = ""
               End If
            Else
               MsgBox "La división no existe", vbOKOnly, "ATENCION"
               Me.txt_division = ""
               Me.txt_division_descripcion = ""
               Me.txt_subdivision = ""
               Me.txt_subdivision_descripcion = ""
               Me.txt_estampado = ""
               Me.txt_estampado_descripcion = ""
               rsaux7.Close
            End If
         Else
            Me.txt_tipo = ""
            Me.txt_tipo_descripcion = ""
            Me.txt_division = ""
            Me.txt_division_descripcion = ""
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            rsaux7.Close
            MsgBox "El tipo de producto no existe", vbOKOnly, "ATENCION"
         End If
      Else
         rsaux9.Open "select * from tb_Articulos where substring(vcha_art_articulo_id,7,5) = '" + Trim(Me.txt_codigo_vianney) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            var_linea = Trim(IIf(IsNull(rsaux9!vcha_lin_linea_id), "", rsaux9!vcha_lin_linea_id))
            var_DEscripcion = IIf(IsNull(rsaux9!vcha_Art_nombre_español), "", rsaux9!vcha_Art_nombre_español)
            var_proveedor = ""
            If var_linea = "00" Then
               linea_textilera = "13"
            End If
            If var_linea = "2" Then
               linea_textilera = "30"
            End If
            If var_linea = "10" Then
               linea_textilera = "12"
            End If
            If var_linea = "11" Then
               linea_textilera = "75"
            End If
            If var_linea = "12" Then
               linea_textilera = "10"
            End If
            If var_linea = "13" Then
               linea_textilera = "30"
            End If
            If var_linea = "14" Then
               linea_textilera = "40"
            End If
            If var_linea = "15" Then
               linea_textilera = "50"
            End If
            If var_linea = "16" Then
               linea_textilera = "20"
            End If
            If var_linea = "20" Then
               linea_textilera = "16"
            End If
            If var_linea = "22" Then
               linea_textilera = "16"
            End If
            If var_linea = "23" Then
               linea_textilera = "16"
            End If
            If var_linea = "24" Then
               linea_textilera = "16"
            End If
            If var_linea = "28" Then
               linea_textilera = "13"
            End If
            If var_linea = "29" Then
               linea_textilera = "12"
            End If
            If var_linea = "30" Then
               linea_textilera = "13"
            End If
            If var_linea = "31" Then
               linea_textilera = "13"
            End If
            If var_linea = "35" Then
               linea_textilera = "16"
            End If
            If var_linea = "39" Then
               linea_textilera = "13"
            End If
            If var_linea = "40" Then
               linea_textilera = "14"
            End If
            If var_linea = "41" Then
               linea_textilera = "14"
            End If
            If var_linea = "42" Then
               linea_textilera = "15"
            End If
            If var_linea = "43" Then
               linea_textilera = "15"
            End If
            If var_linea = "44" Then
               linea_textilera = "25"
            End If
            If var_linea = "45" Then
               linea_textilera = "24"
            End If
            If var_linea = "50" Then
               linea_textilera = "15"
            End If
            If var_linea = "55" Then
               linea_textilera = "13"
            End If
            If var_linea = "59" Then
               linea_textilera = "13"
            End If
            If var_linea = "60" Then
               linea_textilera = "14"
            End If
            If var_linea = "65" Then
               linea_textilera = "13"
            End If
            If var_linea = "70" Then
               linea_textilera = "16"
            End If
            If var_linea = "75" Then
               linea_textilera = "13"
            End If
            If var_linea = "80" Then
               linea_textilera = "16"
            End If
            If var_linea = "90" Then
               linea_textilera = "16"
            End If
            If var_linea = "91" Then
               linea_textilera = "16"
            End If
            If var_linea = "92" Then
               linea_textilera = "16"
            End If
            If var_linea = "93" Then
               linea_textilera = "16"
            End If
            If var_linea = "94" Then
               linea_textilera = "13"
            End If
            If var_linea = "95" Then
               linea_textilera = "13"
            End If
            Me.txt_tipo = "6"
            Me.txt_subdivision = "00"
            Me.txt_division = linea_textilera
            Me.txt_estampado = Me.txt_codigo_vianney
            var_codigo_textilera = "6" + linea_textilera + Me.txt_subdivision + Me.txt_estampado
                     
                     
            rsaux7.Open "select * from tb_tipos_productos where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
            If Not rsaux7.EOF Then
               Me.txt_tipo_descripcion = IIf(IsNull(rsaux7!vcha_tpr_nombre), "", rsaux7!vcha_tpr_nombre)
               rsaux7.Close
               rsaux7.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_div_DIVISION_ID = '" + Me.txt_division + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  Me.txt_division_descripcion = IIf(IsNull(rsaux7!vcha_div_nombre), "", rsaux7!vcha_div_nombre)
                  rsaux7.Close
                  rsaux7.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' AND VCHA_SUB_SUBDIVISION_ID = '" + Me.txt_subdivision + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     Me.txt_subdivision_descripcion = IIf(IsNull(rsaux7!vcha_sub_nombre), "", rsaux7!vcha_sub_nombre)
                     rsaux7.Close
                     rsaux7.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_eST_eSTAMPADO_ID = '" + Me.txt_estampado + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                     If Not rsaux7.EOF Then
                        Me.txt_estampado_descripcion = IIf(IsNull(rsaux7!vcha_est_nombre), "", rsaux7!vcha_est_nombre)
                        rsaux7.Close
                     Else
                        rsaux7.Close
                        rsaux7.Open "insert into tb_estampados (vcha_Est_estampado_id, vcha_est_nombre) values ('" + Me.txt_codigo_vianney + "','" + var_DEscripcion + "')", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
                     End If
                  Else
                     Me.txt_subdivision = ""
                     Me.txt_subdivision_descripcion = ""
                     Me.txt_estampado = ""
                     Me.txt_estampado_descripcion = ""
                     rsaux7.Close
                     MsgBox "La subdivisión no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                  Me.txt_division = ""
                  Me.txt_division_descripcion = ""
                  Me.txt_subdivision = ""
                  Me.txt_subdivision_descripcion = ""
                  Me.txt_estampado = ""
                  Me.txt_estampado_descripcion = ""
                  rsaux7.Close
                  MsgBox "La división no existe", vbOKOnly, "ATENCION"
               End If
            Else
               Me.txt_tipo = ""
               Me.txt_tipo_descripcion = ""
               Me.txt_division = ""
               Me.txt_division_descripcion = ""
               Me.txt_subdivision = ""
               Me.txt_subdivision_descripcion = ""
               Me.txt_estampado = ""
               Me.txt_estampado_descripcion = ""
               rsaux7.Close
               MsgBox "El tipo de producto no existe", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rsaux9.Close
      End If
      rsaux8.Close
      Me.frm_codigo_vianney.Visible = False
      Me.txt_Descuento.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_codigo_vianney.Visible = False
   End If
End Sub

Private Sub txt_codigo_vianney_LostFocus()
   Me.frm_codigo_vianney.Visible = False
End Sub

Private Sub txt_Descuento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_estampado.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_Cantidad.SetFocus
   End If
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_division_Change()
   If Len(Me.txt_division) = 2 Then
      Me.txt_subdivision.SetFocus
   End If
End Sub

Private Sub txt_division_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_codigo_vianney = ""
      Me.frm_codigo_vianney.Visible = True
      Me.txt_codigo_vianney.SetFocus
   End If
   If KeyCode = 37 Then
      Me.txt_tipo.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_subdivision.SetFocus
   End If
   If KeyCode = 113 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT * from TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' order by vcha_DIV_nombre", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_div_division_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "DIVISIONES"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3600
      Else
         lv_lista.ColumnHeaders(2).Width = 3800
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_division_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_division_LostFocus()
   If Trim(txt_division) <> "" Then
      If Trim(txt_tipo) <> "" Then
         If Len(Trim(txt_division)) = 1 Then
            txt_division = "0" + Trim(txt_division)
         End If
         rs.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_division_descripcion = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_Descuento = ""
         Else
            MsgBox "Clave de división no existe", vbOKOnly, "ATENCION"
            Me.txt_division = ""
            Me.txt_division_descripcion = ""
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_Descuento = ""
         End If
         rs.Close
      Else
         MsgBox "No se a indicado un tipo de producto", vbOKOnly, "ATENCION"
         Me.txt_division = ""
         Me.txt_division_descripcion = ""
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_Descuento = ""
      End If
   End If
End Sub

Private Sub txt_estampado_Change()
   If Len(Me.txt_estampado) = 5 Then
      Me.txt_Descuento.SetFocus
   End If
End Sub

Private Sub txt_estampado_GotFocus()
    Me.frm_estampados.Visible = False
End Sub

Private Sub txt_estampado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_codigo_vianney = ""
      Me.frm_codigo_vianney.Visible = True
      Me.txt_codigo_vianney.SetFocus
   End If
   If KeyCode = 37 Then
      Me.txt_subdivision.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_Descuento.SetFocus
   End If
   If KeyCode = 113 Then
   
      Me.lv_estampados.ListItems.Clear
      Me.txt_clave_estampado = ""
      Me.frm_estampados.Visible = True
      Me.txt_clave_estampado.SetFocus
   End If
End Sub

Private Sub txt_estampado_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estampado_LostFocus()
   If Trim(txt_tipo) <> "" Then
      If Trim(txt_division) <> "" Then
         If Trim(txt_subdivision) <> "" Then
            If Trim(Me.txt_estampado) <> "" Then
               If Len(Trim(txt_estampado)) = 1 Then
                  txt_estampado = "0000" + Trim(txt_estampado)
               Else
                 If Len(Trim(txt_estampado)) = 2 Then
                    txt_estampado = "000" + Trim(txt_estampado)
                 Else
                    If Len(Trim(txt_estampado)) = 3 Then
                       txt_estampado = "00" + Trim(txt_estampado)
                    Else
                       If Len(Trim(txt_estampado)) = 4 Then
                          txt_estampado = "0" + Trim(txt_estampado)
                        End If
                     End If
                  End If
               End If
               rs.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_EST_eSTAMPADO_ID = '" + Me.txt_estampado + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_estampado_descripcion = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               Else
                  MsgBox "Clave de estampado no existe", vbOKOnly, "ATENCION"
                  Me.txt_estampado = ""
                  Me.txt_estampado_descripcion = ""
                  Me.txt_Descuento = ""
               End If
               rs.Close
            End If
         Else
            MsgBox "No se a seleccionado una subdivisión", vbOKOnly, "ATENCION"
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_Descuento = ""
         End If
      Else
         MsgBox "No se a seleccionado una división", vbOKOnly, "ATENCION"
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_Descuento = ""
      End If
   Else
      MsgBox "No se a seleccionado un tipo de producto", vbOKOnly, "ATENCION"
      Me.txt_division = ""
      Me.txt_division_descripcion = ""
      Me.txt_subdivision = ""
      Me.txt_subdivision_descripcion = ""
      Me.txt_estampado = ""
      Me.txt_estampado_descripcion = ""
      Me.txt_Descuento = ""
   End If
End Sub

Private Sub txt_subdivision_Change()
   If Len(Me.txt_subdivision) = 2 Then
      Me.txt_estampado.SetFocus
   End If
End Sub

Private Sub txt_subdivision_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_codigo_vianney = ""
      Me.frm_codigo_vianney.Visible = True
      Me.txt_codigo_vianney.SetFocus
   End If
   If KeyCode = 37 Then
      Me.txt_division.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_estampado.SetFocus
   End If
   If KeyCode = 113 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT  * from TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' order by vcha_SUB_nombre", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBDIVISIONES"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3600
      Else
         lv_lista.ColumnHeaders(2).Width = 3800
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_subdivision_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_subdivision_LostFocus()
   If Trim(txt_subdivision) <> "" Then
      If Trim(txt_tipo) <> "" Then
        If Trim(txt_division) <> "" Then
            If Trim(txt_subdivision) <> "" Then
               If Len(Trim(Me.txt_subdivision)) = 1 Then
                  Me.txt_subdivision = "0" + Trim(txt_subdivision)
               End If
               rs.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_SUB_SUBDIVISION_ID = '" + Me.txt_subdivision + "' AND VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_subdivision_descripcion = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
               Else
                  MsgBox "Clave de subdivisión incorrecta", vbOKOnly, "ATENCION"
                  Me.txt_subdivision = ""
                  Me.txt_subdivision_descripcion = ""
                  Me.txt_estampado = ""
                  Me.txt_estampado_descripcion = ""
                  Me.txt_Descuento = ""
               End If
               rs.Close
            Else
               Me.txt_subdivision = ""
               Me.txt_subdivision_descripcion = ""
               Me.txt_estampado = ""
               Me.txt_estampado_descripcion = ""
               Me.txt_Descuento = ""
            End If
         Else
            MsgBox "No se a indicado una división", vbOKOnly, "ATENCION"
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_Descuento = ""
         End If
      Else
         MsgBox "No se a indicado un tipo de producto", vbOKOnly, "ATENCION"
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_Descuento = ""
      End If
   End If
End Sub

Private Sub txt_tipo_Change()
   If Len(Me.txt_tipo) = 1 Then
      Me.txt_division.SetFocus
   End If
End Sub

Private Sub txt_tipo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 39 Then
      Me.txt_division.SetFocus
   End If
   If KeyCode = 113 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOS_PRODUCTOS order by vcha_tpr_nombre", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_tpr_tipo_producto_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPOS PRODUCTOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3600
      Else
         lv_lista.ColumnHeaders(2).Width = 3800
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 116 Then
      Me.txt_codigo_vianney = ""
      Me.frm_codigo_vianney.Visible = True
      Me.txt_codigo_vianney.SetFocus
   End If
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
   
End Sub

Private Sub txt_tipo_LostFocus()
   If Trim(txt_tipo) = "" Then
      Me.txt_tipo_descripcion = ""
      Me.txt_division = ""
      Me.txt_division_descripcion = ""
      Me.txt_subdivision = ""
      Me.txt_subdivision_descripcion = ""
      Me.txt_estampado = ""
      Me.txt_estampado_descripcion = ""
      Me.txt_Descuento = ""
   Else
      'Me.txt_tipo = "6"
      rs.Open "SELECT * FROM TB_TIPOS_PRODUCTOS WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_tipo_descripcion = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
         Me.txt_division = ""
         Me.txt_division_descripcion = ""
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_Descuento = ""
      Else
         MsgBox "Tipo de producto no existe", vbOKOnly, "ATENCION"
         Me.txt_tipo_descripcion = ""
         Me.txt_division = ""
         Me.txt_division_descripcion = ""
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_Descuento = ""
      End If
      rs.Close
   End If
End Sub

