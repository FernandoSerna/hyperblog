VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmarchivo_envios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_alta_codigo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmarchivo_envios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Insertar codigos de textilera"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_articulo 
      Height          =   915
      Left            =   3240
      TabIndex        =   10
      Top             =   2010
      Width           =   2475
      Begin VB.TextBox txt_articulo 
         Height          =   390
         Left            =   90
         TabIndex        =   11
         Top             =   345
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Artículo"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   12
         Top             =   15
         Width           =   2460
      End
   End
   Begin VB.Frame frm_archivo 
      Height          =   915
      Left            =   150
      TabIndex        =   6
      Top             =   465
      Width           =   2475
      Begin VB.TextBox txt_archivo 
         Height          =   390
         Left            =   90
         TabIndex        =   8
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Nota o archivo"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmarchivo_envios.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Alt + B"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9030
      Picture         =   "frmarchivo_envios.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmarchivo_envios.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   5835
      Left            =   105
      TabIndex        =   1
      Top             =   405
      Width           =   9330
      Begin MSComctlLib.ListView lv_archivo 
         Height          =   5250
         Left            =   45
         TabIndex        =   2
         Top             =   150
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   9260
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nota"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Externo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Interno"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   6262
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Proveedor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "catalogo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "año"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   45
         Top             =   2505
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":0988
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":1262
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":1B3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":20D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":29B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":328E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":3B68
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":3C7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":3D8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":3E9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmarchivo_envios.frx":3FB0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
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
         Height          =   330
         Left            =   5625
         TabIndex        =   9
         Top             =   5445
         Width           =   3510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   105
      TabIndex        =   0
      Top             =   270
      Width           =   9345
   End
End
Attribute VB_Name = "frmarchivo_envios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()

End Sub

Private Sub cmd_aceptar_Click()
   var_si = 1
   For var_i = 1 To Me.lv_archivo.ListItems.Count
       lv_archivo.ListItems.Item(var_i).Selected = True
       If Trim(lv_archivo.selectedItem.SubItems(2)) = "" Then
          var_si = 0
       End If
   Next var_i
   If var_si = 1 Then
      var_si = MsgBox("Se cargara el archivo en el sistema", vbYesNo, "ATENCION")
      If var_si = 6 Then
         If var_empresa = "18" Then
            If var_clave_usuario_global = "U0000000074" Then
               var_cadena = "select * from tb_archivo_comparacion where vcha_com_referencia = 'EPT" + Trim(lv_archivo.selectedItem) + "'"
            Else
               var_cadena = "select * from tb_archivo_comparacion where vcha_com_referencia = 'ERV" + Trim(lv_archivo.selectedItem) + "'"
            End If
         Else
            var_cadena = "select * from tb_archivo_comparacion where vcha_com_referencia = 'EPTE" + Trim(lv_archivo.selectedItem) + "'"
         End If
         rsaux5.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If rsaux5.EOF Then
            Set TB_ARCH_COMPARACION_I = New TB_ARCH_COMPARACION_I
            Set TB_ARCHIVO_COMPARACION_I_ENVIOS = New TB_ARCHIVO_COMPARACION_I_ENVIOS
            For var_i = 1 To lv_archivo.ListItems.Count
                lv_archivo.ListItems.Item(var_i).Selected = True
                If var_empresa = "18" Then
                   If var_clave_usuario_global = "U0000000142" Or var_clave_usuario_global = "U0000000105" Then
                      ok = TB_ARCHIVO_COMPARACION_I_ENVIOS.Anadir(var_empresa, var_unidad_organizacional, "PTTEX", "EPV", lv_archivo.selectedItem, Date, "U", lv_archivo.selectedItem.SubItems(5), lv_archivo.selectedItem.SubItems(2), lv_archivo.selectedItem.SubItems(6), lv_archivo.selectedItem.SubItems(4), 0, "", "EPT" + Trim(lv_archivo.selectedItem), 0, 0, 2005, "", 0)
                   Else
                      ok = TB_ARCHIVO_COMPARACION_I_ENVIOS.Anadir(var_empresa, var_unidad_organizacional, "RVTEX", "ERV", lv_archivo.selectedItem, Date, "U", lv_archivo.selectedItem.SubItems(5), lv_archivo.selectedItem.SubItems(2), lv_archivo.selectedItem.SubItems(6), lv_archivo.selectedItem.SubItems(4), 0, "", "ERV" + Trim(lv_archivo.selectedItem), 0, 0, 2005, "", 0)
                   End If
                Else
                   ok = TB_ARCH_COMPARACION_I.Anadir(var_empresa, var_unidad_organizacional, "8", "EPTE", lv_archivo.selectedItem, Date, "U", lv_archivo.selectedItem.SubItems(5), lv_archivo.selectedItem.SubItems(2), lv_archivo.selectedItem.SubItems(6), lv_archivo.selectedItem.SubItems(4), 0, "", "EPTE" + Trim(lv_archivo.selectedItem), 0, 0, 2005, "", 0)
                End If
            Next var_i
            MsgBox "Se a terminado la carga del archivo", vbOKOnly, "ATENCION"
         Else
            MsgBox "El archivo ya fue cargado", vbOKOnly, "ATENCION"
         End If
         rsaux5.Close
      End If
   Else
      MsgBox "Faltan artículos con código", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_alta_codigo_Click()
Dim VERIFICADOR As Integer
Dim var_codigo_textilera As String
   If var_empresa = "18" Then
      If Me.lv_archivo.ListItems.Count > 0 Then
         For var_j = 1 To lv_archivo.ListItems.Count
             lv_archivo.ListItems.Item(var_j).Selected = True
             If Trim(lv_archivo.selectedItem.SubItems(2)) = "" Then
                var_equivalencia = Trim(lv_archivo.selectedItem.SubItems(1))
                var_linea = Trim(lv_archivo.selectedItem.SubItems(8))
                var_costo = lv_archivo.selectedItem.SubItems(6)
                var_precio = lv_archivo.selectedItem.SubItems(7)
                var_catalogo = Trim(lv_archivo.selectedItem.SubItems(9))
                var_DEscripcion = Trim(lv_archivo.selectedItem.SubItems(3))
                var_proveedor = Trim(lv_archivo.selectedItem.SubItems(5))
                var_nota = Trim(lv_archivo.selectedItem)
                var_año = Trim(lv_archivo.selectedItem.SubItems(10))
                If var_linea = "00" Then
                   linea_textilera = "13"
                End If
                If var_linea = "10" Then
                   linea_textilera = "12"
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
                var_codigo = Mid(var_equivalencia, 7, 5)
                var_codigo_textilera = "6" + var_linea + "00" + var_codigo + "0"
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
                 
                rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_aRTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_cadena = "INSERT INTO TB_ARTICULOS (VCHA_aRT_aRTICULO_ID, VCHA_aRT_nombre_español, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR, VCHA_LIN_LINEA_ID, VCHA_ART_CATALOGO_VIGENTE, VCHA_TPR_TIPO_PRODUCTO_ID,       VCHA_DIV_DIVISION_ID,            VCHA_SUB_SUBDIVISION_ID,          VCHA_EST_ESTAMPADO_ID) VALUES"
                   var_cadena = var_cadena + "('" + var_codigo_textilera + "', '" + var_DEscripcion + "', " + CStr(var_precio) + ", " + CStr(var_costo) + ",    '" + var_linea + "',  '" + var_catalogo + "', '" + Mid(var_codigo_textilera, 1, 1) + "','" + Mid(var_codigo_textilera, 2, 2) + "', '" + Mid(var_codigo_textilera, 4, 2) + "', '" + Mid(var_codigo_textilera, 6, 5) + "')"
                   rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                   rsaux.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('01','" + var_codigo_textilera + "', " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                   rsaux.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Mid(var_codigo_textilera, 6, 5) + "'"
                   If rsaux.EOF Then
                      rsaux2.Open "insert into tb_estampados (vcha_est_estampado_id, vcha_est_nombre) values ('" + Mid(var_codigo_textilera, 6, 5) + "', '" + var_DEscripcion + "')", cnn, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux.Close
                   rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn, adOpenDynamic, adLockOptimistic
                   If rsaux.EOF Then
                      rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id) values ('" + var_equivalencia + "', '" + var_codigo_textilera + "')", cnn, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux.Close
                End If
                rs.Close
                var_cadena = "update TB_ARCHIVOS_ENVIOS set vcha_Art_articulo_id = '" + var_codigo_textilera + "' where vcha_aco_proveedor = '" + var_proveedor + "' and inte_aco_numero = " + CStr(var_nota) + " and vcha_aco_codigo_Externo = '" + var_equivalencia + "' and INTE_ACO_AÑO = " + CStr(var_año)
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                lv_archivo.selectedItem.SubItems(2) = var_codigo_textilera
                var_codigo_1 = Mid(var_codigo_textilera, 1, 10)
                var_codigo_general = var_codigo_textilera
                rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + var_codigo_textilera + "' and vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn, adOpenDynamic, adLockOptimistic
                If rsaux.EOF Then
                   rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_articulo_id) values ('" + var_equivalencia + "', '" + var_codigo_textilera + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
                rsaux.Close
                
                For var_i = 0 To 9
                    var_codigo = var_codigo_1 + Trim(CStr(var_i))
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
                
                    rs.Open "SELECT * from tb_reclasificacion where vcha_alm_almacen_id = 'RVTEX' and vcha_art_articulo_id = '" + var_codigo + "' and vcha_rec_codigo_general = '" + var_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
                    If rs.EOF Then
                       rsaux3.Open "INSERT INTO TB_RECLASIFICACION (VCHA_ALM_ALMACEN_ID, vcha_Art_articulo_id, vcha_rec_codigo_general) values ('RVTEX','" + var_codigo + "','" + var_codigo_general + "')", cnn, adOpenDynamic, adLockOptimistic
                    End If
                    rs.Close
                    
                    rs.Open "SELECT * from tb_reclasificacion where vcha_alm_almacen_id = 'RETEX' and vcha_art_articulo_id = '" + var_codigo + "' and vcha_rec_codigo_general = '" + var_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
                    If rs.EOF Then
                       rsaux3.Open "INSERT INTO TB_RECLASIFICACION (VCHA_ALM_ALMACEN_ID, vcha_Art_articulo_id, vcha_rec_codigo_general) values ('RETEX','" + var_codigo + "','" + var_codigo_general + "')", cnn, adOpenDynamic, adLockOptimistic
                    End If
                    rs.Close
                Next var_i
             End If
         Next var_j
      End If
   Else
      If Me.lv_archivo.ListItems.Count > 0 Then
         rs.Open "select * from tb_archivos_envios where inte_aco_numero = " + Me.lv_archivo.selectedItem, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  var_equivalencia = (rs!vcha_aco_codigo_externo)
                  rsaux1.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + var_equivalencia + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     rsaux3.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + IIf(IsNull(rsaux1!vcha_Art_Articulo_id), "", rsaux1!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        rsaux2.Open "update tb_archivos_envios set vcha_Art_Articulo_id = '" + IIf(IsNull(rsaux1!vcha_Art_Articulo_id), "", rsaux1!vcha_Art_Articulo_id) + "' where inte_aco_consecutivo = " + CStr(rs!inte_aco_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux2.Open "update tb_archivos_envios set vcha_Art_Articulo_id = '' where inte_aco_consecutivo = " + CStr(rs!inte_aco_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux3.Close
                  Else
                     rsaux2.Open "update tb_archivos_envios set vcha_Art_Articulo_id = '' where inte_aco_consecutivo = " + CStr(rs!inte_aco_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  rs.MoveNext
            Wend
         End If
         rs.Close
         
         
         
         
         
         rs.Open "select * from TB_ARCHIVOS_ENVIOS where inte_aco_numero = " + Me.lv_archivo.selectedItem, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lv_archivo.ListItems.Clear
            Dim var_cantidad As Double
            var_cantidad = 0
            While Not rs.EOF
                  Set list_item = lv_archivo.ListItems.Add(, , rs!inte_aco_numero)
                  list_item.SubItems(1) = rs!vcha_aco_codigo_externo
                  list_item.SubItems(3) = rs!vcha_aco_descripcion_externa
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id)
                  list_item.SubItems(4) = rs!floa_Aco_Cantidad
                  list_item.SubItems(5) = rs!vcha_aco_proveedor
                  list_item.SubItems(6) = rs!floa_aco_costo
                  list_item.SubItems(7) = IIf(IsNull(rs!floa_Aco_precio), 0, rs!floa_Aco_precio)
                  list_item.SubItems(8) = IIf(IsNull(rs!vcha_lin_linea_id), "", rs!vcha_lin_linea_id)
                  list_item.SubItems(9) = IIf(IsNull(rs!vcha_cat_catalogo_id), "", rs!vcha_cat_catalogo_id)
                  list_item.SubItems(10) = IIf(IsNull(rs!INTE_ACO_AÑO), 2005, rs!INTE_ACO_AÑO)
                  var_cantidad = var_cantidad + rs!floa_Aco_Cantidad
                  rs.MoveNext
            Wend
            Me.lbl_cantidad = Format(var_cantidad, "###,###,##0.00")
         Else
            MsgBox "Número de nota incorrecto", vbOKOnly, "ATENCION"
         End If
         rs.Close
         
         
         
         
         
         
         
         
         
      Else
         MsgBox "No se a seleccionado una nota", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_buscar_Click()
   frm_archivo.Visible = True
   Me.txt_archivo = ""
   Me.txt_archivo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 500
   Left = 1400
   frm_archivo.Visible = False
   frm_Articulo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_archivo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_Articulo.Visible = True
      txt_articulo = ""
      txt_articulo.SetFocus
   End If
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   Dim list_item As ListItem
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      rs.Open "select * from TB_ARCHIVOS_ENVIOS where inte_aco_numero = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         lv_archivo.ListItems.Clear
         Dim var_cantidad As Double
         var_cantidad = 0
         While Not rs.EOF
            Set list_item = lv_archivo.ListItems.Add(, , rs!inte_aco_numero)
            list_item.SubItems(1) = rs!vcha_aco_codigo_externo
            list_item.SubItems(3) = rs!vcha_aco_descripcion_externa
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id)
            list_item.SubItems(4) = rs!floa_Aco_Cantidad
            list_item.SubItems(5) = rs!vcha_aco_proveedor
            list_item.SubItems(6) = rs!floa_aco_costo
            list_item.SubItems(7) = IIf(IsNull(rs!floa_Aco_precio), 0, rs!floa_Aco_precio)
            list_item.SubItems(8) = IIf(IsNull(rs!vcha_lin_linea_id), "", rs!vcha_lin_linea_id)
            list_item.SubItems(9) = IIf(IsNull(rs!vcha_cat_catalogo_id), "", rs!vcha_cat_catalogo_id)
            list_item.SubItems(10) = IIf(IsNull(rs!INTE_ACO_AÑO), 2005, rs!INTE_ACO_AÑO)
            var_cantidad = var_cantidad + rs!floa_Aco_Cantidad
            rs.MoveNext
         Wend
         Me.lbl_cantidad = Format(var_cantidad, "###,###,##0.00")
      Else
         MsgBox "Número de nota incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
      frm_archivo.Visible = False
   End If
End Sub

Private Sub txt_archivo_LostFocus()
   frm_archivo.Visible = False
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Dim list_item As ListItem
      If Trim(txt_articulo) <> "" Then
         rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_si = MsgBox("¿Desea aplicar el codigo?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar la aplicación del artículo", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rsaux.Open "UPDATE TB_ARCHIVOS_ENVIOS SET VCHA_ACO_DESCRIPCION_EXTERNA = '" + rs!vcha_Art_nombre_español + "', VCHA_ART_ARTICULO_ID = '" + txt_articulo + "' WHERE INTE_ACO_NUMERO = '" + lv_archivo.selectedItem + "' AND VCHA_ACO_CODIGO_EXTERNO = '" + lv_archivo.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_archivo = lv_archivo.selectedItem
                  lv_archivo.ListItems.Clear
                  rsaux.Open "select * from TB_ARCHIVOS_ENVIOS where inte_aco_numero = " + Me.txt_archivo, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Dim var_cantidad As Double
                     var_cantidad = 0
                     While Not rsaux.EOF
                           Set list_item = lv_archivo.ListItems.Add(, , rsaux!inte_aco_numero)
                           list_item.SubItems(1) = rsaux!vcha_aco_codigo_externo
                           list_item.SubItems(3) = rsaux!vcha_aco_descripcion_externa
                           list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id)
                           list_item.SubItems(4) = rsaux!floa_Aco_Cantidad
                           list_item.SubItems(5) = rsaux!vcha_aco_proveedor
                           list_item.SubItems(6) = rsaux!floa_aco_costo
                           var_cantidad = var_cantidad + rsaux!floa_Aco_Cantidad
                           rsaux.MoveNext
                     Wend
                     Me.lbl_cantidad = Format(var_cantidad, "###,###,##0.00")
                  Else
                     MsgBox "Número de nota incorrecto", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
                  
               End If
            End If
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      frm_Articulo.Visible = False
   End If
End Sub

Private Sub txt_articulo_LostFocus()
   frm_Articulo.Visible = False
End Sub
