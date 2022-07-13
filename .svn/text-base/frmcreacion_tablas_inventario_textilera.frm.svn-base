VERSION 5.00
Begin VB.Form frmcreacion_tablas_inventario_textilera 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   150
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2775
      Width           =   4500
   End
   Begin VB.CommandButton cmd_crear_tablas 
      Caption         =   "Crear Tablas"
      Height          =   2505
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   4485
   End
End
Attribute VB_Name = "frmcreacion_tablas_inventario_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer

Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub cmd_crear_tablas_Click()
   rs.Open "set excl on", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "delete from articulos", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "pack", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "select distinct substring(tb_articulos.vcha_Art_articulo_id,1,10) as codigo, cast(substring(tb_articulos.vcha_Art_articulo_id,1,1) as integer) as tipo, cast(substring(tb_articulos.vcha_art_articulo_id,2,2) as integer) as division, cast(substring(tb_articulos.vcha_art_articulo_id,4,2) as integer) as subdivisio, cast(substring(tb_articulos.vcha_Art_Articulo_id,6,5) as integer) as estampado, isnull(inte_art_salida_masiva,0) as recontable from tb_Articulos, tb_existencias, tb_almacenes where tb_articulos.vcha_art_articulo_id = tb_existencias.vcha_art_articulo_id and tb_almacenes.vcha_emp_empresa_id ='18' and tb_Existencias.vcha_alm_almacen_id = tb_almacenes.vcha_alm_almacen_id and len(tb_articulos.vcha_Art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into articulos (codigo, tipo,division, subdivisio, estampado, recontable, costo, costo_prov, preciou, preciour,cve_unidad, maximo, minimo, bulto, posicion, reclasific, codigorecl,decoradora, oferta, cve_linea, materia_p, avios, mano_o, gastos_f, cve_catalo) values ('" + rs!codigo + "'," + CStr(rs!tipo) + "," + CStr(rs!division) + "," + CStr(rs!subdivisio) + "," + CStr(rs!estampado) + "," + CStr(rs!recontable) + ",0,0,0,0,1,0,0,0,'','','',0, 0,'',0,0,0,0, '')", var_tabla, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "set excl on", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tipoprod", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "pack", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "select cast(vcha_tpr_tipo_producto_id as integer) as cve_produc, vcha_tpr_nombre as descripcio from tb_tipos_productos", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into tipoprod (cve_produc, descripcio) values (" + CStr(rs!cve_produc) + ",'" + rs!descripcio + "')", var_tabla, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "set excl on", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "delete from division", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "pack", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "select cast(vcha_div_division_id as integer) as cve_divisi, vcha_div_nombre as descripcio, cast(vcha_tpr_tipo_producto_id as integer) as depende from tb_divisiones", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into division (cve_divisi, descripcio, depende) values (" + CStr(rs!cve_divisi) + ",'" + rs!descripcio + "'," + CStr(rs!depende) + " )", var_tabla, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "set excl on", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "delete from subdivis", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "pack", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "select cast(vcha_sub_subdivision_id as integer) as cve_divi, vcha_sub_nombre as descripcio, cast(vcha_tpr_tipo_producto_id as integer) as depende, cast(vcha_div_division_id as integer) as depende2 from tb_subdivisiones", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Text1 = "insert into (cve_divi, descripcio, depende, depende2) values (" + CStr(rs!cve_divi) + ",'" + Trim(rs!descripcio) + "'," + CStr(rs!depende) + "," + CStr(rs!depende2) + ")"
         rsaux.Open "insert into subdivis (cve_divi, descripcio, depende, depende2) values (" + CStr(rs!cve_divi) + ",'" + Trim(rs!descripcio) + "'," + CStr(rs!depende) + "," + CStr(rs!depende2) + ")", var_tabla, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "set excl on", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "delete from estampad", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "pack", var_tabla, adOpenDynamic, adLockOptimistic
   rs.Open "select cast(vcha_Est_estampado_id as integer) as cve_esta, vcha_est_nombre as descripcio from tb_estampados where len(vcha_est_estampado_id) <= 5", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_cadena = "insert into estampad (cve_Esta, descripcio,depende) values (" + CStr(rs!cve_esta) + ",'" + rs!descripcio + "',0)"
         'MsgBox var_cadena, vbOKOnly, ""
         
         
         
         
         
         
         
         
         
         rsaux.Open "insert into estampad (cve_Esta, descripcio,depende) values (" + CStr(rs!cve_esta) + ",'" + rs!descripcio + "',0)", var_tabla, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Load()
   Set var_tabla = CreateObject("ADODB.connection")
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=c:\sistemas\desarrollo\recuperacion;SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub
