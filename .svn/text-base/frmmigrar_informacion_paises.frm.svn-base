VERSION 5.00
Begin VB.Form frmmigrar_informacion_paises 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Agrupadores"
      Height          =   585
      Left            =   165
      TabIndex        =   9
      Top             =   4620
      Width           =   3120
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Lista precios exportaciones"
      Height          =   525
      Left            =   165
      TabIndex        =   8
      Top             =   4005
      Width           =   3120
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ubicaciones"
      Height          =   435
      Left            =   150
      TabIndex        =   7
      Top             =   3510
      Width           =   3090
   End
   Begin VB.Frame Frame1 
      Caption         =   "CALCULA VERIFICADOR "
      Height          =   1185
      Left            =   45
      TabIndex        =   4
      Top             =   2235
      Width           =   3285
      Begin VB.TextBox txt_verificador 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   765
         Width           =   2760
      End
      Begin VB.TextBox txt_codigo 
         Height          =   345
         Left            =   165
         MaxLength       =   11
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   750
      Left            =   405
      TabIndex        =   3
      Top             =   1410
      Width           =   2580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ciudades de usa"
      Height          =   660
      Left            =   390
      TabIndex        =   2
      Top             =   675
      Width           =   2610
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Catálogos"
      Height          =   570
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Artículos"
      Height          =   570
      Left            =   405
      TabIndex        =   0
      Top             =   90
      Width           =   1275
   End
End
Attribute VB_Name = "frmmigrar_informacion_paises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection

Private Sub Command1_Click()
   Dim mcodigo As String
   Dim sum1 As Double
   Dim sum2 As Double
   Dim verificador As Integer
   Dim icont As Integer
   Dim msuma As Integer
   Dim var_fecha As Date
   Dim dia As Integer, mes As Integer, año As Integer
   rs.Open "delete from tb_Articulos", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_equivalencias", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select cveestilo, descripcio, fechaalta, costo, preciolist, linea, familia from estilos", var_tabla, adOpenDynamic, adLockOptimistic
   MsgBox CStr(Date), vbOKOnly, ""
   If Not rs.EOF Then
      While Not rs.EOF
            mcodigo = "646244" + IIf(IsNull(rs!cveestilo), "", rs!cveestilo)
            sum1 = 0
            sum2 = 0
            For icont = 1 To 11
                If ((icont / 2) - Int((icont / 2))) = 0 Then
                   sum2 = sum2 + CInt(Mid(mcodigo, icont, 1))
                Else
                   sum1 = sum1 + CInt(Mid(mcodigo, icont, 1))
                End If
            Next icont
            msuma = sum1 * 13 + sum2
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
            mcodigo = Trim(mcodigo) + Trim(CStr(verificador))
            var_fecha_string = Format(IIf(IsNull(rs!fechaalta), "", rs!fechaalta), "Short Date")
            var_fecha = Format(IIf(IsNull(rs!fechaalta), "", rs!fechaalta), "Short Date")
            dia = Day(rs!fechaalta)
            mes = Month(rs!fechaalta)
            año = Year(rs!fechaalta)
            If var_fecha_string = "30/12/1899" Then
               var_fecha = Date
            Else
               var_fecha = rs!fechaalta
            End If
            rsaux.Open "insert into tb_articulos (vcha_art_articulo_id, VCHA_ART_NOMBRE_ESPAÑOL, MONE_ART_PRECIO_BASE, MONE_ART_COSTO_ESTANDAR, DTIM_ART_FECHA_ALTA, VCHA_ART_CATALOGO_VIGENTE, VCHA_LIN_LINEA_ID) values ('" + mcodigo + "', '" + rs!descripcio + "', " + CStr(rs!preciolist) + ", " + CStr(rs!costo) + ",'" + CStr(var_fecha) + "','" + rs!familia + "','" + rs!linea + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_equivalencias (vcha_art_articulo_id, VCHA_EQU_CODIGO_EQUIVALENTE) values ('" + mcodigo + "', '" + Trim(rs!cveestilo) + "')", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
   MsgBox "ya termino", vbOKOnly, "atencion"
End Sub

Private Sub Command2_Click()
   rs.Open "delete from tb_catalogos", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select cvefamilia, nombre, cveagrcata from familias", var_tabla, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            rsaux.Open "insert into tb_Catalogos (vcha_cat_catalogo_id, vcha_cat_nombre, vcha_agr_agrupador_catalogo_id) values ('" + rs!cvefamilia + "','" + rs!nombre + "', '" + rs!cveagrcata + "')", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
   End If
   rs.Close
End Sub

Private Sub Command3_Click()
   rs.Open "select * from codigos_usa", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into tb_colonias (vcha_pai_pais_id, vcha_est_estado_id, vcha_mun_municipio_id, vcha_ciu_ciudad_id, vcha_col_colonia_id, vcha_col_nombre, vcha_col_cp) values ('002','" + rs!clave_est + "','" + rs!clave_mun + "', '" + rs!clave_ciud + "', '" + rs!clave_col + "', '','" + rs!codigo + "' )", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command4_Click()
   Dim var_codigo As String, mcodigo As String, var_clave_lista As String
   rs.Open "select * from tb_detalle_lista_precios", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            mcodigo = rs!VCHA_ART_ARTICULO_ID
            var_codigo = rs!VCHA_ART_ARTICULO_ID
            var_clave_lista = rs!vcha_lis_lista_precios_id
            sum1 = 0
            sum2 = 0
            For icont = 1 To 11
                If ((icont / 2) - Int((icont / 2))) = 0 Then
                   sum2 = sum2 + CInt(Mid(mcodigo, icont, 1))
                Else
                   sum1 = sum1 + CInt(Mid(mcodigo, icont, 1))
                End If
            Next icont
            msuma = sum1 * 13 + sum2
            verificador = Int(10 - ((msuma / 10) - Int(msuma / 10)) * 10)
            If verificador = 10 Then
               verificador = 0
            End If
            mcodigo = Trim(mcodigo) + Trim(CStr(verificador))
            rsaux.Open "update tb_detalle_lista_precios set vcha_art_articulo_id = '" + mcodigo + "' where vcha_art_articulo_id = '" + var_codigo + "' and vcha_lis_lista_precios_id = '" + var_clave_lista + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
       Wend
   End If
End Sub

Private Sub Command5_Click()
   rs.Open "select cveestilo, descripcio, fechaalta, costo, preciolist, linea, familia, ubicacion from estilos", var_tabla, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rs!cveestilo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux2.Open "insert into tb_ubicaciones_almacen (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_ubi_ubicacion_1) values ('8','" + rsaux!VCHA_ART_ARTICULO_ID + "', '" + rs!ubicacion + "')", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
            rs.MoveNext
       Wend
   End If
   rs.Close
End Sub

Private Sub Command6_Click()
   rs.Open "select cveestilo, base from precios", var_tabla, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!cveestilo) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux2.Open "insert into tb_detalle_lista_precios (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) values ('04', '" + rsaux!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!base) + ")", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.MoveNext
      Wend
   End If
   rs.Close
End Sub

Private Sub Command7_Click()
    rs.Open "delete from tb_agrupadores", cnn, adOpenDynamic, adLockOptimistic
    rs.Open "delete from TB_DETALLE_AGRUPADORES", cnn, adOpenDynamic, adLockOptimistic
    rs.Open "select distinct nomempingl from estilos where allt(nomempingl) <> ''", var_tabla, adOpenDynamic, adLockOptimistic
    var_i = 1
    While Not rs.EOF
          If rsaux.State = 1 Then
             rsaux.Close
          End If
          If Trim(rs!nomempingl) <> "" Then
             rsaux.Open "insert into tb_agrupadores (VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_AGR_NOMBRE, VCHA_AGR_TIPO) values ('1','" + CStr(var_i) + "', '" + rs!nomempingl + "','1')", cnn, adOpenDynamic, adLockOptimistic
          End If
          var_i = var_i + 1
          rs.MoveNext
    Wend
    If rs.State = 1 Then
       rs.Close
    End If
    rs.Open "select cveestilo, nomempingl from estilos where allt(nomempingl) <> ''", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!cveestilo) + "'", cnn, adOpenDynamic, adLockOptimistic
          var_articulo = ""
          If Not rsaux.EOF Then
             var_articulo = rsaux!VCHA_ART_ARTICULO_ID
          End If
          rsaux.Close
          rsaux.Open "select * from tb_agrupadores where VCHA_AGR_NOMBRE = '" + Trim(rs!nomempingl) + "'", cnn, adOpenDynamic, adLockOptimistic
          var_agrupador = ""
          If Not rsaux.EOF Then
             var_agrupador = rsaux!vcha_agr_agrupador_id
          End If
          rsaux.Close
          rsaux4.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO,VCHA_ART_ARTICULO_ID,VCHA_LIN_LINEA_ID,VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values ('" + var_agrupador + "',1,'" + var_articulo + "','','','','')", cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
    
    
    
    rs.Open "select distinct nomexporta from estilos where allt(nomexporta) <> ''", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          If rsaux.State = 1 Then
             rsaux.Close
          End If
          If Trim(rs!nomexporta) <> "" Then
             rsaux.Open "insert into tb_agrupadores (VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_AGR_NOMBRE, VCHA_AGR_TIPO) values ('2','" + CStr(var_i) + "', '" + rs!nomexporta + "','1')", cnn, adOpenDynamic, adLockOptimistic
          End If
          var_i = var_i + 1
          rs.MoveNext
    Wend
    If rs.State = 1 Then
       rs.Close
    End If
    rs.Open "select cveestilo, nomexporta from estilos where allt(nomexporta) <> ''", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!cveestilo) + "'", cnn, adOpenDynamic, adLockOptimistic
          var_articulo = ""
          If Not rsaux.EOF Then
             var_articulo = rsaux!VCHA_ART_ARTICULO_ID
          End If
          rsaux.Close
          rsaux.Open "select * from tb_agrupadores where VCHA_AGR_NOMBRE = '" + Trim(rs!nomexporta) + "' and vcha_fag_familia_agrupador_id = '2'", cnn, adOpenDynamic, adLockOptimistic
          var_agrupador = ""
          If Not rsaux.EOF Then
             var_agrupador = rsaux!vcha_agr_agrupador_id
          End If
          rsaux.Close
          rsaux4.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO,VCHA_ART_ARTICULO_ID,VCHA_LIN_LINEA_ID,VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values ('" + var_agrupador + "',1,'" + var_articulo + "','','','','')", cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
    
End Sub

Private Sub Form_Load()
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=c:\sistemas\desarrollo\integral\;SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
End Sub


Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      Dim verificador As Integer
      mcodigo = txt_codigo
      sum1 = 0
      sum2 = 0
      For icont = 1 To 11
          If ((icont / 2) - Int((icont / 2))) = 0 Then
             sum2 = sum2 + CInt(Mid(mcodigo, icont, 1))
          Else
             sum1 = sum1 + CInt(Mid(mcodigo, icont, 1))
          End If
      Next icont
      msuma = sum1 * 13 + sum2
      verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
      If verificador = 10 Then
         verificador = 0
      End If
      txt_verificador = Trim(mcodigo) + Trim(CStr(verificador))
   End If
End Sub
