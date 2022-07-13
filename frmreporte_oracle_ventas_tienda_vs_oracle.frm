VERSION 5.00
Begin VB.Form frmreporte_oracle_ventas_tienda_vs_oracle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación Tienda VS Oracle"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_notas_credito 
      Appearance      =   0  'Flat
      Caption         =   "D"
      Height          =   315
      Left            =   435
      Picture         =   "frmreporte_oracle_ventas_tienda_vs_oracle.frx":0000
      TabIndex        =   8
      ToolTipText     =   "Devoluciones"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   105
      TabIndex        =   2
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3990
      Picture         =   "frmreporte_oracle_ventas_tienda_vs_oracle.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Caption         =   "F"
      Height          =   315
      Left            =   105
      Picture         =   "frmreporte_oracle_ventas_tienda_vs_oracle.frx":073C
      TabIndex        =   0
      ToolTipText     =   "Facturación"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   15
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_oracle_ventas_tienda_vs_oracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_mes As Integer

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         If CDate(Me.txt_fecha_inicio) <= CDate(Me.txt_fecha_fin) Then
            var_fecha_fin_1 = CDate(Me.txt_fecha_fin) + 1
            var_dia = CStr(Day(CDate(Me.txt_fecha_inicio)))
            var_mes = CStr(Month(CDate(txt_fecha_inicio)))
            var_año = CStr(Year(CDate(txt_fecha_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
           
            var_dia = CStr(Day(CDate(var_fecha_fin_1)))
            var_mes = CStr(Month(CDate(var_fecha_fin_1)))
            var_año = CStr(Year(CDate(var_fecha_fin_1)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT VCHA_ART_TIPO_CANTIA, substring(a.VCHA_ART_ARTICULO_ID,1,5) as vcha_art_articulo_id, min(VCHA_ART_NOMBRE_ESPAÑOL) as VCHA_ART_NOMBRE_ESPAÑOL , sum(FLOA_SAL_CANTIDAD) as CANTIDAD FROM TB_SALIDAS a, TB_ARTICULOS b WHERE DTIM_SAL_FECHA >= " + var_fecha_inicio + " and DTIM_SAL_FECHA  < " + var_fecha_fin + " and a.VCHA_ART_ARTICULO_ID= b.VCHA_ART_ARTICULO_ID and b.VCHA_ART_TIPO_CANTIA in ('T','M') AND VCHA_MOV_MOVIMIENTO_ID IN ('CC_2','FA', 'FV','VDI') group by VCHA_ART_TIPO_CANTIA, substring(a.VCHA_ART_ARTICULO_ID,1,5)"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "select max(inte_tem_consecutivo) from TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_cadena_oracle_1 = ""
               var_cadena_oracle_2 = ""
               var_cadena_oracle_3 = ""
               var_cadena_oracle_4 = ""
               var_j = 0
               While Not rs.EOF
                     var_codigo_cantia = Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5)
                     var_codigo_oracle = ""
                     If var_j <= 500 Then
                        If var_cadena_oracle_1 = "" Then
                           var_cadena_oracle_1 = "'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                        Else
                           var_cadena_oracle_1 = var_cadena_oracle_1 + ",'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                        End If
                     Else
                         If var_j > 500 And var_j <= 1000 Then
                            If var_cadena_oracle_2 = "" Then
                               var_cadena_oracle_2 = "'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                            Else
                               var_cadena_oracle_2 = var_cadena_oracle_2 + ",'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                            End If
                         Else
                            If var_j > 1000 And var_j <= 1500 Then
                               If var_cadena_oracle_3 = "" Then
                                  var_cadena_oracle_3 = "'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                               Else
                                  var_cadena_oracle_3 = var_cadena_oracle_3 + ",'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                               End If
                            Else
                               If var_j > 1500 Then
                                  If var_cadena_oracle_4 = "" Then
                                     var_cadena_oracle_4 = "'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                                  Else
                                     var_cadena_oracle_4 = var_cadena_oracle_4 + ",'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                                  End If
                               End If
                            End If
                         End If
                     End If
                     var_cadena = "insert into TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE (inte_tem_consecutivo, fecha_inicio, fecha_fin, vcha_art_Articulo_id, vcha_art_nombre_español, cantidad_tienda, cantidad_oracle, codigo_oracle)  values (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "-1,'" + CStr(rs!VCHA_aRT_ARTICULO_ID) + "', '" + rs!vcha_Art_nombre_español + "', " + CStr(rs!Cantidad) + ", 0, '" + var_codigo_oracle + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     var_j = var_j + 1
                     rs.MoveNext
               Wend
               If var_cadena_oracle_1 <> "" Then
                  rsaux1.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, CROSS_REFERENCE FROM mtl_cross_references_v A, MTL_SYSTEM_ITEMS_B B Where a.inventory_item_id = B.inventory_item_id  AND CROSS_REFERENCE in (" + var_cadena_oracle_1 + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set codigo_oracle = '" + rsaux1!segment1 + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_art_articulo_id = '" + rsaux1!CROSS_REFERENCE + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
               End If
               If var_cadena_oracle_2 <> "" Then
                  rsaux1.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, CROSS_REFERENCE FROM mtl_cross_references_v A, MTL_SYSTEM_ITEMS_B B Where a.inventory_item_id = B.inventory_item_id  AND CROSS_REFERENCE in (" + var_cadena_oracle_2 + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set codigo_oracle = '" + rsaux1!segment1 + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_art_articulo_id = '" + rsaux1!CROSS_REFERENCE + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
               End If
               
               If var_cadena_oracle_3 <> "" Then
                  rsaux1.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, CROSS_REFERENCE FROM mtl_cross_references_v A, MTL_SYSTEM_ITEMS_B B Where a.inventory_item_id = B.inventory_item_id  AND CROSS_REFERENCE in (" + var_cadena_oracle_3 + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set codigo_oracle = '" + rsaux1!segment1 + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_art_articulo_id = '" + rsaux1!CROSS_REFERENCE + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
               End If
               
               If var_cadena_oracle_4 <> "" Then
                  rsaux1.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, CROSS_REFERENCE FROM mtl_cross_references_v A, MTL_SYSTEM_ITEMS_B B Where a.inventory_item_id = B.inventory_item_id  AND CROSS_REFERENCE in (" + var_cadena_oracle_4 + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set codigo_oracle = '" + rsaux1!segment1 + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_art_articulo_id = '" + rsaux1!CROSS_REFERENCE + "'", cnn, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
               End If
               
               
               rsaux1.Open "SELECT c.segment1, b.item_description, sum(shipped_quantity) as cantidad FROM OE_ORDER_HEADERS_ALL a, wsh_deliverables_v b, xxvia_system_items_b c WHERE b.inventory_item_id = c.inventory_item_id and  b.organization_id = c.organization_id and a.orig_sys_document_ref like 'SIDCAN%' and a.header_id = b.SOURCE_HEADER_id and a.ordered_Date >= TO_DATE('" + Me.txt_fecha_inicio + "','DD/MM/YYYY') and a.ordered_date < TO_DATE('" + CStr(CDate(Me.txt_fecha_fin) + 1) + "','DD/MM/YYYY') and released_status = 'C' group by c.segment1, b.item_description"
               While Not rsaux1.EOF
                     rsaux.Open "select * from TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo_oracle = '" + rsaux1!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set cantidad_oracle = " + CStr(rsaux1!Cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo_oracle = '" + rsaux1!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux2.Open "insert into TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE (inte_tem_consecutivo, fecha_inicio, fecha_fin, vcha_art_articulo_id, vcha_art_nombre_español, cantidad_tienda, cantidad_oracle, codigo_oracle) values  (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "-1,'','" + rsaux1!item_description + "',0," + CStr(rsaux1!Cantidad) + ", '" + rsaux1!segment1 + "')", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "delete from TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_ventas_cantia_tienda_vs_oracle.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VENTAS_TIENDA_VS_ORACLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Catálogo de artículos por almacén"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_ventas_cantia_tienda_vs_oracle.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_VENTAS_TIENDA_VS_ORACLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\facturacion_cantia_tienda_vs_oracle_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
               rsaux.Open "delete from TB_TEMP_ORACLE_COMPARACION_ENVIOS_CONSIGNACION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               
            Else
               MsgBox "No existen artículos facturados para el periodo seleccionado", vbo
            End If
            rs.Close
            
         Else
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inciio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub



Private Sub cmd_nuevo_Click()

End Sub

   
Private Sub cmd_notas_credito_Click()

   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         If CDate(Me.txt_fecha_inicio) <= CDate(Me.txt_fecha_fin) Then
            var_fecha_fin_1 = CDate(Me.txt_fecha_fin) + 1
            var_dia = CStr(Day(CDate(Me.txt_fecha_inicio)))
            var_mes = CStr(Month(CDate(txt_fecha_inicio)))
            var_año = CStr(Year(CDate(txt_fecha_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
           
            var_dia = CStr(Day(CDate(var_fecha_fin_1)))
            var_mes = CStr(Month(CDate(var_fecha_fin_1)))
            var_año = CStr(Year(CDate(var_fecha_fin_1)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT VCHA_ART_TIPO_CANTIA, substring(a.VCHA_ART_ARTICULO_ID,1,5) as vcha_art_articulo_id, min(VCHA_ART_NOMBRE_ESPAÑOL) as VCHA_ART_NOMBRE_ESPAÑOL , sum(FLOA_ent_CANTIDAD) as CANTIDAD FROM TB_ENTRADAS  a, TB_ARTICULOS b WHERE DTIM_ent_FECHA >= " + var_fecha_inicio + " and DTIM_ent_FECHA  < " + var_fecha_fin + " and a.VCHA_ART_ARTICULO_ID= b.VCHA_ART_ARTICULO_ID and b.VCHA_ART_TIPO_CANTIA in ('T','M') AND VCHA_MOV_MOVIMIENTO_ID IN ('CC_4','DC') group by VCHA_ART_TIPO_CANTIA, substring(a.VCHA_ART_ARTICULO_ID,1,5)"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "select max(inte_tem_consecutivo) from TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_cadena_oracle = ""
               While Not rs.EOF
                     var_codigo_cantia = Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5)
                     var_codigo_oracle = ""
                     If var_cadena_oracle = "" Then
                        var_cadena_oracle = "'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                     Else
                        var_cadena_oracle = var_cadena_oracle + ",'" + Mid(rs!VCHA_aRT_ARTICULO_ID, 1, 5) + "'"
                     End If
                     var_cadena = "insert into TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE (inte_tem_consecutivo, fecha_inicio, fecha_fin, vcha_art_Articulo_id, vcha_art_nombre_español, cantidad_tienda, cantidad_oracle, codigo_oracle)  values (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "-1,'" + CStr(rs!VCHA_aRT_ARTICULO_ID) + "', '" + rs!vcha_Art_nombre_español + "', " + CStr(rs!Cantidad) + ", 0, '" + var_codigo_oracle + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, CROSS_REFERENCE FROM mtl_cross_references_v A, MTL_SYSTEM_ITEMS_B B Where a.inventory_item_id = B.inventory_item_id  AND CROSS_REFERENCE in (" + var_cadena_oracle + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set codigo_oracle = '" + rsaux1!segment1 + "' where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_art_articulo_id = '" + rsaux1!CROSS_REFERENCE + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "SELECT C.SEGMENT1, C.DESCRIPTION AS item_description, SUM(ORDERED_QUANTITY) AS CANTIDAD FROM OE_ORDER_HEADERS_ALL A, OE_ORDER_LINES_ALL B, XXVIA_SYSTEM_ITEMS_B C WHERE A.ORIG_SYS_DOCUMENT_REF LIKE 'SIDDCCAN%' AND A.HEADER_ID = B.HEADER_ID AND a.ordered_Date >= TO_DATE('" + Me.txt_fecha_inicio + "','DD/MM/YYYY') and a.ordered_date < TO_DATE('" + CStr(CDate(Me.txt_fecha_fin) + 1) + "','DD/MM/YYYY') AND A.SHIP_FROM_ORG_ID = c.organization_id AND B.INVENTORY_ITEM_ID = C.INVENTORY_ITEM_ID GROUP BY C.SEGMENT1, C.DESCRIPTION", cnnoracle_4, adOpenDynamic, adLockOptimistic
               'rsaux1.Open "SELECT c.segment1, b.item_description, sum(shipped_quantity) as cantidad FROM OE_ORDER_HEADERS_ALL a, wsh_deliverables_v b, xxvia_system_items_b c WHERE b.inventory_item_id = c.inventory_item_id and  b.organization_id = c.organization_id and a.orig_sys_document_ref like 'SIDCAN%' and a.header_id = b.SOURCE_HEADER_id and a.ordered_Date >= TO_DATE('" + Me.txt_fecha_inicio + "','DD/MM/YYYY') and a.ordered_date < TO_DATE('" + CStr(CDate(Me.txt_fecha_fin) + 1) + "','DD/MM/YYYY') and released_status = 'C' group by c.segment1, b.item_description"
               While Not rsaux1.EOF
                     rsaux.Open "select * from TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo_oracle = '" + rsaux1!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        rsaux2.Open "update TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE set cantidad_oracle = " + CStr(rsaux1!Cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo_oracle = '" + rsaux1!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux2.Open "insert into TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE (inte_tem_consecutivo, fecha_inicio, fecha_fin, vcha_art_articulo_id, vcha_art_nombre_español, cantidad_tienda, cantidad_oracle, codigo_oracle) values  (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "-1,'','" + rsaux1!item_description + "',0," + CStr(rsaux1!Cantidad) + ", '" + rsaux1!segment1 + "')", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               rsaux1.Open "delete from TB_TEMP_ORACLE_VENTAS_TIENDA_VS_ORACLE where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_devoluciones_cantia_tienda_vs_oracle.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VENTAS_TIENDA_VS_ORACLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Catálogo de artículos por almacén"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_devoluciones_cantia_tienda_vs_oracle.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_VENTAS_TIENDA_VS_ORACLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\devoluciones_cantia_tienda_vs_oracle_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
               rsaux.Open "delete from TB_TEMP_ORACLE_COMPARACION_ENVIOS_CONSIGNACION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               
            Else
               MsgBox "No existen artículos facturados para el periodo seleccionado", vbo
            End If
            rs.Close
            
         Else
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inciio incorrecta", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2900
   Left = 3850
   txt_fecha_inicio = Date
   txt_fecha_fin = Date
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_reporte_acumulado_ventas)
End Sub


Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_fecha_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_fin) Then
         frmcalendario.mes = CDate(Me.txt_fecha_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_inicio) Then
         frmcalendario.mes = CDate(Me.txt_fecha_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha_inicio = var_fecha_general
   End If
End Sub


