VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfactura_orden_compra_estampados 
   BorderStyle     =   0  'None
   Caption         =   "Factura de Estampados el Refugio"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   " Artículo "
      Height          =   990
      Left            =   150
      TabIndex        =   7
      Top             =   1080
      Width           =   7350
      Begin VB.Label lbl_articulo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   105
         TabIndex        =   8
         Top             =   210
         Width           =   7170
      End
   End
   Begin VB.Frame frm_codigo 
      Height          =   30
      Left            =   1245
      TabIndex        =   4
      Top             =   5955
      Width           =   1830
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Text            =   "0000000000000"
         Top             =   495
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Código Interno"
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   6
         Top             =   330
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4050
      Left            =   135
      TabIndex        =   2
      Top             =   2010
      Width           =   7365
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   3870
         Left            =   60
         TabIndex        =   3
         Top             =   135
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   6826
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
            Text            =   "Código Externo"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Costo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Facura "
      Height          =   660
      Left            =   150
      TabIndex        =   0
      Top             =   405
      Width           =   7350
      Begin VB.TextBox txt_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2730
         TabIndex        =   1
         Top             =   225
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6195
      Left            =   0
      TabIndex        =   9
      Top             =   -60
      Width           =   7620
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Seleccione el artículo"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   15
         TabIndex        =   10
         Top             =   105
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmfactura_orden_compra_estampados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   var_posible = 0
   If lv_articulos.ListItems.Count > 0 Then
      
   End If
   rs.Open "SEelct * from tb_archivo_comparacion where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = 'PTTEX' AND VCHA_MOV_MOVIMIENTO_ID = 'EC' AND INTE_COM_NUMERO = " + Me.txt_factura, cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      MsgBox "La factura ya fue cargada anteriormente", vbOKOnly, "ATENCION"
   Else
      Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
      ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, "PTTEX", "EPV", lv_archivo.selectedItem, Date, "U", lv_archivo.selectedItem.SubItems(5), lv_archivo.selectedItem.SubItems(2), lv_archivo.selectedItem.SubItems(6), lv_archivo.selectedItem.SubItems(4), 0, "", "EPT" + Trim(lv_archivo.selectedItem), 0, 0, 2005, "", 0)
      
   End If
   rs.Close
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.txt_factura = frmentradas.txt_factura
   frm_codigo.Visible = False
   If Trim(Me.txt_factura) <> "" Then
      rsaux.Open "SELECT distinct * FROM TB_FACTURAS_ESTAMPADOS_REFUGIO WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "'", cnn, adOpenDynamic, adLockOptimistic
      'MsgBox Me.txt_factura
      If rsaux.EOF Then
         'cnn_importacion.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=VENTAS;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
         cnn_importacion.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDTEXTILERA;Data Source=sqlquezada2"
         cnn_importacion.CursorLocation = adUseClient
         'var_cadena = "SELECT DISTINCT cfaven.faccod, lfaven.FACSER, lfaven.facdsc descripcion, lfaven.facmts cantidad, (lfaven.facpremts + albrec.albrpre) precio,  albrec.albrpre tela, lfaven.faclin From cfaven@cipic.vianney.com.mx, lfaven@cipic.vianney.com.mx, barcad@cipic.vianney.com.mx, barpie@cipic.vianney.com.mx, albrec@cipic.vianney.com.mx, calprd@cipic.vianney.com.mx, clienv@cipic.vianney.com.mx Where cfaven.emprcod = lfaven.emprcod AND cfaven.faccod = lfaven.faccod AND cfaven.emprcod = clienv.emprcod"
         'var_cadena = var_cadena + " AND cfaven.clicod = clienv.clicod AND lfaven.emprcod = barcad.emprcod AND lfaven.facbarcod = barcad.barcod aND lfaven.facbarreo = barcad.barcodreo AND lfaven.facbarpar = barcad.barcodpar AND lfaven.emprcod = calprd.emprcod AND lfaven.facalbcod = calprd.albprocod AND barcad.emprcod = barpie.emprcod aND barcad.barcod = barpie.barcod"
         'var_cadena = var_cadena + " AND barcad.barcodreo = barpie.barcodreo AND barcad.barcodpar = barpie.barcodpar AND barpie.emprcod = albrec.emprcod AND barpie.albreccod = albrec.albreccod AND calprd.albdomenv = clienv.clienvlin AND cfaven.faccod = " + Me.txt_factura
         var_cadena = "SELECT INTE_CAR_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_COSTO+floa_sal_precio as floa_Sal_costo, VCHA_SAL_DESCRIPCION_FACTURA, FLOA_SAL_CANTIDAD From dbo.TB_SALIDAS WHERE     (VCHA_EMP_EMPRESA_ID = '15') AND (VCHA_CAR_DOCUMENTO = 'FA') AND (INTE_CAR_NUMERO = " + Me.txt_factura + ")"
         rs.Open var_cadena, cnn_importacion, adOpenDynamic, adLockOptimistic
         
         var_consecutivo = 0
         While Not rs.EOF
               var_consecutivo = var_consecutivo + 1
               Set list_item = lv_articulos.ListItems.Add(, , Trim(rs!vcha_Art_Articulo_id))
               list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura))
               list_item.SubItems(2) = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
               list_item.SubItems(3) = IIf(IsNull(rs!floa_Sal_costo), 0, rs!floa_Sal_costo)
               list_item.SubItems(4) = var_consecutivo
               rsaux2.Open "insert into tb_facturas_estampados_refugio (vcha_fac_factura, vcha_fac_codigo_externo, vcha_art_articulo_id, vcha_fac_descripcion, floa_fac_cantidad, floa_fac_COSTO, inte_fac_consecutivo) values ('" + Me.txt_factura + "', '" + CStr(rs!facser) + "','', '" + rs!descripcion + "'," + CStr(rs!Cantidad) + "," + CStr(rs!Precio) + ", " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         cnn_importacion.Close
      Else
         While Not rsaux.EOF
               Set list_item = lv_articulos.ListItems.Add(, , Trim(rsaux!vcha_fac_codigo_externo))
               list_item.SubItems(1) = Trim(IIf(IsNull(rsaux!vcha_fac_descripcion), "", rsaux!vcha_fac_descripcion))
               list_item.SubItems(2) = IIf(IsNull(rsaux!floa_fac_cantidad), 0, rsaux!floa_fac_cantidad)
               list_item.SubItems(3) = IIf(IsNull(rsaux!floa_fac_costo), 0, rsaux!floa_fac_costo)
               list_item.SubItems(4) = IIf(IsNull(rsaux!inte_fac_consecutivo), 0, rsaux!inte_fac_consecutivo)
               rsaux.MoveNext
         Wend
      End If
      rsaux.Close
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
   frmentradas.txt_codigo = var_codigo_tela

End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then
       var_si = MsgBox("¿Desea seleccionar el costo del artículo?", vbYesNo, "ATENCION")
       If var_si = 6 Then
          If lv_articulos.ListItems.Count > 0 Then
             var_costo_tela = Me.lv_articulos.selectedItem.SubItems(3)
          Else
             var_costo_tela = 0
          End If
          Unload Me
       End If
    End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_codigo) <> "" Then
         rs.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "UPDATE TB_FACTURAS_ESTAMPADOS_REFUGIO SET VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "' WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "' AND VCHA_FAC_CODIGO_EXTERNO = '" + Me.lv_articulos.selectedItem + "' and inte_fac_consecutivo = " + Me.lv_articulos.selectedItem.SubItems(3), cnn, adOpenDynamic, adLockOptimistic
            lv_articulos.selectedItem.SubItems(1) = Me.txt_codigo
            lv_articulos.SetFocus
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
   If KeyAscii = 27 Then
      frm_codigo.Visible = False
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   Me.frm_codigo.Visible = False
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_factura_LostFocus()
'9940721
'9156687
   If Trim(Me.txt_factura) <> "" Then
      rsaux.Open "SELECT * FROM TB_FACTURAS_ESTAMPADOS_REFUGIO WHERE VCHA_FAC_FACTURA = '" + Me.txt_factura + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux.EOF Then
         'cnn_importacion.Open "Provider=OraOLEDB.Oracle.1;User ID=ACATEX;Data Source=ORCL;Extended Properties=;Persist Security Info=True;Password=ACATEX"
         cnn_importacion.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=VENTAS;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
         
         cnn_importacion.CursorLocation = adUseClient
         var_cadena = "SELECT DISTINCT cfaven.faccod, lfaven.FACSER, lfaven.facdsc descripcion, lfaven.facmts cantidad, (lfaven.facpremts + albrec.albrpre) precio,  albrec.albrpre tela, lfaven.faclin From cfaven@cipic.vianney.com.mx, lfaven@cipic.vianney.com.mx, barcad@cipic.vianney.com.mx, barpie@cipic.vianney.com.mx, albrec@cipic.vianney.com.mx, calprd@cipic.vianney.com.mx, clienv@cipic.vianney.com.mx Where cfaven.emprcod = lfaven.emprcod AND cfaven.faccod = lfaven.faccod AND cfaven.emprcod = clienv.emprcod"
         var_cadena = var_cadena + " AND cfaven.clicod = clienv.clicod AND lfaven.emprcod = barcad.emprcod AND lfaven.facbarcod = barcad.barcod aND lfaven.facbarreo = barcad.barcodreo AND lfaven.facbarpar = barcad.barcodpar AND lfaven.emprcod = calprd.emprcod AND lfaven.facalbcod = calprd.albprocod AND barcad.emprcod = barpie.emprcod aND barcad.barcod = barpie.barcod"
         var_cadena = var_cadena + " AND barcad.barcodreo = barpie.barcodreo AND barcad.barcodpar = barpie.barcodpar AND barpie.emprcod = albrec.emprcod AND barpie.albreccod = albrec.albreccod AND calprd.albdomenv = clienv.clienvlin AND cfaven.faccod = " + Me.txt_factura
         rs.Open var_cadena, cnn_importacion, adOpenDynamic, adLockOptimistic
         var_consecutivo = 0
         While Not rs.EOF
               var_consecutivo = var_consecutivo + 1
               Set list_item = lv_articulos.ListItems.Add(, , Trim(rs!facser))
               list_item.SubItems(1) = ""
               list_item.SubItems(2) = Trim(IIf(IsNull(rs!descripcion), "", rs!descripcion))
               list_item.SubItems(3) = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
               list_item.SubItems(4) = var_consecutivo
               rsaux2.Open "insert into tb_facturas_estampados_refugio (vcha_fac_factura, vcha_fac_codigo_externo, vcha_art_articulo_id, vcha_fac_descripcion, floa_fac_cantidad, floa_fac_COSTO, inte_fac_consecutivo) values ('" + Me.txt_factura + "', '" + CStr(rs!facser) + "','', '" + rs!descripcion + "'," + CStr(rs!Cantidad) + "," + CStr(rs!Precio) + ", " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         cnn_importacion.Close
      Else
         While Not rsaux.EOF
               Set list_item = lv_articulos.ListItems.Add(, , Trim(rsaux!vcha_fac_codigo_externo))
               list_item.SubItems(1) = ""
               list_item.SubItems(2) = Trim(IIf(IsNull(rsaux!vcha_fac_descripcion), "", rsaux!vcha_fac_descripcion))
               list_item.SubItems(3) = IIf(IsNull(rsaux!floa_fac_cantidad), 0, rsaux!floa_fac_cantidad)
               list_item.SubItems(4) = IIf(IsNull(rsaux!inte_fac_consecutivo), 0, rsaux!inte_fac_consecutivo)
               rsaux.MoveNext
         Wend
      End If
      rsaux.Close
   End If
End Sub
