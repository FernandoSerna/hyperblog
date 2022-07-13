VERSION 5.00
Begin VB.Form frmgenera_archivo_distribucion_coppel_EXCEL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar archivo coppel excel"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Pedido "
      Height          =   1575
      Left            =   90
      TabIndex        =   6
      Top             =   435
      Width           =   3405
      Begin VB.TextBox txt_pedido 
         Height          =   350
         Left            =   1665
         TabIndex        =   2
         Top             =   315
         Width           =   1260
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   350
         Left            =   1665
         TabIndex        =   3
         Top             =   690
         Width           =   1260
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   350
         Left            =   1665
         TabIndex        =   4
         Top             =   1065
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido de COPPEL:"
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   390
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicio:"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha fin:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   75
      TabIndex        =   5
      Top             =   360
      Width           =   3510
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      Picture         =   "frmgenera_archivo_distribucion_coppel_EXCEL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_generar_archivo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmgenera_archivo_distribucion_coppel_EXCEL.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Generar Archivo"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmgenera_archivo_distribucion_coppel_EXCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub cmd_generar_archivo_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         var_fecha_fin_1 = CDate(txt_fecha_fin) + 1
         var_dia = CStr(Day(CDate(txt_fecha_inicio)))
         var_mes = CStr(Month(CDate(txt_fecha_inicio)))
         var_año = CStr(Year(CDate(txt_fecha_inicio)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
          
          
         var_dia = CStr(Day(var_fecha_fin_1))
         var_mes = CStr(Month(var_fecha_fin_1))
         var_año = CStr(Year(var_fecha_fin_1))
         If Len(Trim(var_dia)) = 1 Then
             var_dia = "0" + var_dia
          End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         
         Set var_tabla = CreateObject("ADODB.connection")
         Dim var_pedido As String
         var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=c:\coppel\archivos\;SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
         
         rs.Open "select * from tb_pedido_original_coppel where archivo = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "DELETE FROM TB_aRCHIVO_DISTRIBUCION_COPPEL where archivo = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "SELECT DISTINCT INTE_CAR_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_sER_SERIE_ID, INTE_EMB_EMBARQUE, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_IMPORTE_IVA, destino FROM VW_ARCHIVO_DISTRIBUCION_COPPEL WHERE archivo = '" + Me.txt_pedido + "' and DTIM_CAR_FECHA >= " + var_fecha_inicio + " and dtim_Car_fecha <= " + var_fecha_fin + "-0.000001", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_orden_surtido = rsaux!INTE_ORS_ORDEN_SURTIDO
                  VAR_EMBARQUE = rsaux!inte_emb_embarque
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "SELECT MAX(INTE_PAQ_CAJA) FROM TB_DETALLE_CAJAS WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido) + " AND INTE_EMB_EMBARQUE = " + CStr(VAR_EMBARQUE) + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                  VAR_MAXIMO_cAJA = CStr(rsaux2(0).Value)
                  If Len(VAR_MAXIMO_cAJA) = 1 Then
                     VAR_MAXIMO_cAJA = "00" + VAR_MAXIMO_cAJA
                  Else
                    If Len(VAR_MAXIMO_cAJA) = 2 Then
                       VAR_MAXIMO_cAJA = "0" + VAR_MAXIMO_cAJA
                    End If
                  End If
                  rsaux2.Close
                  VAR_ET_DATO6 = "FACTURA:" + CStr(rsaux!inte_Car_numero)
                  rsaux2.Open "SELECT SUM(FLOA_SAL_CANTIDAD) FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO = " + CStr(rsaux!inte_Car_numero) + " AND VCHA_SER_sERIE_ID = '" + rsaux!VCHA_SER_SERIE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_cantidad = rsaux2(0).Value
                  VAR_ET_DATO7 = "UNIDS.FACTURA:" + CStr(var_cantidad)
                  rsaux2.Close
                  cnn.CommandTimeout = 360
                  rsaux2.Open "SELECT * FROM VW_ARCHIVO_DISTRIBUCION_COPPEL WHERE archivo = '" + Me.txt_pedido + "' and destino = '" + rsaux!Destino + "' and DTIM_CAR_FECHA >= " + var_fecha_inicio + " and dtim_Car_fecha <= " + var_fecha_fin + "-0.000001", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux2.EOF
                        var_año_s = CStr(Year(rsaux2!dtim_Car_fecha))
                        var_mes_s = CStr(Month(rsaux2!dtim_Car_fecha))
                        var_dia_s = CStr(Day(rsaux2!dtim_Car_fecha))
                        If Len(var_mes_s) = 1 Then
                           var_mes_s = "0" + var_mes_s
                        End If
                        If Len(var_dia_s) = 1 Then
                           var_dia_s = "0" + var_dia_s
                        End If
                        var_fecha_s = var_año_s + "-" + var_mes_s + "-" + var_dia_s
                        var_nombre_pedido = rsaux2!NUMPEDIDO + "A"
                        var_caja = CStr(rsaux2!inte_paq_caja)
                        If Len(var_caja) = 1 Then
                           var_caja = "00" + var_caja
                        Else
                           If Len(var_caja) = 2 Then
                              var_caja = "0" + var_caja
                           End If
                        End If
                        VAR_ET_DATO1 = rsaux2!ET_DATO1
                        VAR_ET_DATO2 = rsaux2!ET_DATO2
                        VAR_ET_DATO3 = rsaux2!ET_DATO3
                        VAR_ET_DATO4 = rsaux2!ET_DATO4
                        VAR_ET_DATO5 = rsaux2!ET_DATO5
                        VAR_ET_DATO8 = rsaux2!ET_DATO8
                        VAR_ET_DATO9 = rsaux2!ET_DATO9
                        VAR_ET_DATO10 = var_caja + "/" + VAR_MAXIMO_cAJA
                        VAR_ET_DATO11 = "P" + Trim(Me.txt_pedido) + Trim(rsaux2!Destino) + Trim(var_caja) + Trim(VAR_MAXIMO_cAJA)
                        VAR_ET_DATO12 = "P " + Trim(Me.txt_pedido) + " " + Trim(rsaux2!Destino) + " " + Trim(var_caja) + " " + Trim(VAR_MAXIMO_cAJA)
                        rsaux4.Open "SELECT SUM(FLOA_SAL_CANTIDAD) FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO = " + CStr(rsaux!inte_Car_numero) + " AND VCHA_SER_sERIE_ID = '" + rsaux!VCHA_SER_SERIE_ID + "' AND VCHA_ART_ARTICULO_ID = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_cantidad_surtida = rsaux2!Cantidad
                        rsaux4.Close
                        var_cadena = "INSERT INTO TB_ARCHIVO_DISTRIBUCION_COPPEL (ARCHIVO, BODRECIBO, DESTINO, MODELOPROV, CODICOPPEL, TALLA_COPP, CANTPED, CANTSUR, COSTO, VENTA, LOTE, TOTLOTES, NUMFACTURA, NUMPEDIDO, IMPTEFACTU, IVAFACTURA, UNIDSFACTU, PROVEEDOR, ET_DATO1, ET_DATO2,"
                        var_cadena = var_cadena + "ET_DATO3, ET_DATO4, ET_DATO5, ET_DATO6, ET_DATO7, ET_DATO8, ET_DATO9, ET_DATO10, ET_DATO11, ET_DATO12, CDDESTINO, MARCA, FAMILIA, TRANSF, EMPAQUE, PEDIMENTO, PTO_ENT, PAIS_ORI, IMPORT, NETO, PORCEN, INDICE, COMPLETO, FECHAFAC)"
                        var_cadena = var_cadena + " Values ('" + rsaux2!archivo + "', '" + rsaux2!BODRECIBO + "','" + rsaux2!Destino + "', '" + rsaux2!MODELOPROV + "', '" + rsaux2!codicoppel + "', '" + rsaux2!TALLA_COPP + "', " + CStr(var_cantidad_surtida) + "," + CStr(var_cantidad_surtida) + "," + CStr(rsaux2!Costo) + ", " + CStr(rsaux2!VENTA) + ", '" + var_caja + "', '" + VAR_MAXIMO_cAJA + "', " + CStr(rsaux!inte_Car_numero) + ", '" + rsaux2!NUMPEDIDO + "', " + CStr(rsaux!floa_Car_importe_neto) + ", " + CStr(rsaux!floa_car_importe_iva) + ", " + CStr(var_cantidad) + ", '" + rsaux2!proveedor + "',"
                        var_cadena = var_cadena + "'" + VAR_ET_DATO1 + "', '" + VAR_ET_DATO2 + "', '" + VAR_ET_DATO3 + "', '" + VAR_ET_DATO4 + "', '" + VAR_ET_DATO5 + "', '" + VAR_ET_DATO6 + "', '" + VAR_ET_DATO7 + "', '" + VAR_ET_DATO8 + "', '" + VAR_ET_DATO9 + "', '" + VAR_ET_DATO10 + "', '" + VAR_ET_DATO11 + "', '" + VAR_ET_DATO12 + "', '" + rsaux2!CDDESTINO + "', '" + rsaux2!marca + "', '" + rsaux2!familia + "', '" + rsaux2!TRANSF + "', '" + rsaux2!EMPAQUE + "', '" + rsaux2!PEDIMENTO + "', '" + rsaux2!PTO_ENT + "', '" + rsaux2!PAIS_ORI + "', '" + rsaux2!Import + "', " + CStr(rsaux2!neto) + ", "
                        var_cadena = var_cadena + CStr(rsaux2!PORCEN) + ", " + CStr(rsaux2!indice) + ", '" + rsaux2!COMPLETO + "', null)"
                        rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.MoveNext
                  Wend
                  rsaux.MoveNext
            Wend
            rsaux.Close
            rsaux.Open "SELECT * FROM TB_PEDIDO_ORIGINAL_COPPEL WHERE archivo = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_Archivo = rsaux!archivo
                  var_pedido = rsaux!NUMPEDIDO
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "SELECT * FROM TB_ARCHIVO_DISTRIBUCION_COPPEL WHERE ARCHIVO = '" + rsaux!archivo + "' AND BODRECIBO = '" + rsaux!BODRECIBO + "' AND DESTINO = '" + rsaux!Destino + "' AND MODELOPROV = '" + rsaux!MODELOPROV + "'AND CODICOPPEL = '" + rsaux!codicoppel + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux2.EOF Then
                     var_cadena = "INSERT INTO TB_ARCHIVO_DISTRIBUCION_COPPEL (ARCHIVO, BODRECIBO, DESTINO, MODELOPROV, CODICOPPEL, TALLA_COPP, CANTPED, CANTSUR, COSTO, VENTA, LOTE, TOTLOTES, NUMFACTURA, NUMPEDIDO, IMPTEFACTU, IVAFACTURA, UNIDSFACTU, PROVEEDOR, ET_DATO1, ET_DATO2,"
                     var_cadena = var_cadena + "ET_DATO3, ET_DATO4, ET_DATO5, ET_DATO6, ET_DATO7, ET_DATO8, ET_DATO9, ET_DATO10, ET_DATO11, ET_DATO12, CDDESTINO, MARCA, FAMILIA, TRANSF, EMPAQUE, PEDIMENTO, PTO_ENT, PAIS_ORI, IMPORT, NETO, PORCEN, INDICE, COMPLETO, FECHAFAC)"
                     var_cadena = var_cadena + " Values ('" + rsaux!archivo + "', '" + rsaux!BODRECIBO + "','" + rsaux!Destino + "', '" + rsaux!MODELOPROV + "', '" + rsaux!codicoppel + "', '" + rsaux!TALLA_COPP + "', " + CStr(rsaux!cantped) + ",0," + CStr(rsaux!Costo) + ", " + CStr(rsaux!VENTA) + ", '', '', '', '" + rsaux!NUMPEDIDO + "', 0, 0, 0, '" + rsaux!proveedor + "',"
                     var_cadena = var_cadena + "'" + rsaux!ET_DATO1 + "', '" + rsaux!ET_DATO2 + "', '" + rsaux!ET_DATO3 + "', '" + rsaux!ET_DATO4 + "', '" + rsaux!ET_DATO5 + "', '" + rsaux!ET_DATO6 + "', '" + rsaux!ET_DATO7 + "', '" + rsaux!ET_DATO8 + "', '" + rsaux!ET_DATO9 + "', '" + rsaux!ET_DATO10 + "', '" + rsaux!ET_DATO11 + "', '" + rsaux!ET_DATO12 + "', '" + rsaux!CDDESTINO + "', '" + rsaux!marca + "', '" + rsaux!familia + "', '" + rsaux!TRANSF + "', '" + rsaux!EMPAQUE + "', '" + rsaux!PEDIMENTO + "', '" + rsaux!PTO_ENT + "', '" + rsaux!PAIS_ORI + "', '" + rsaux!Import + "', " + CStr(rsaux!neto) + ", "
                     var_cadena = var_cadena + CStr(rsaux!PORCEN) + ", " + CStr(rsaux!indice) + ", '" + rsaux!COMPLETO + "', NULL)"
                     rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rsaux.MoveNext
            Wend
            rsaux.Close
            
            rsaux.Open "UPDATE TB_ARCHIVO_DISTRIBUCION_COPPEL SET FECHAFAC =  '" + var_fecha_s + "'", cnn, adOpenDynamic, adLockOptimistic
            
         
            Set reporte = appl.OpenReport(App.Path + "\rep_archivo_distribucion_coppel.rpt")
            reporte.RecordSelectionFormula = "{TB_ARCHIVO_DISTRIBUCION_COPPEL.ARCHIVO} = '" + Me.txt_pedido + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\p1" + var_nombre_pedido + ".xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
         
         
         
         
         
         Else
            MsgBox "No existe el pedido", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   Me.txt_fecha_inicio = Date
   Me.txt_fecha_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub txt_fecha_Change()
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_generar_archivo.SetFocus
   End If
End Sub

Private Sub txt_fecha_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_fecha_fin.SetFocus
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      rs.Open "delete from tb_pedido_original_coppel where archivo = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select * from tb_Archivo_pedido_coppel_excel where vcha_arc_pedido = '" + Me.txt_pedido + "' order by numbodegadestino", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            var_cadena = "insert into tb_pedido_original_coppel (archivo,            bodrecibo,                   destino,                      modeloprov,                  codicoppel,                  talla_copp,                  cantped,                   cantsur,               costo,                  venta,                    lote, totlotes, numfactura, numpedido, imptefactu, ivafactura, unidsfactu, proveedor, et_Dato1, et_Dato2, et_dato3, et_dato4, et_dato5, et_dato6, et_dato7, et_Dato8, et_dato9, et_dato10, et_dato11, et_dato12, cddestino, marca, familia, transf, empaque, pedimento, pto_ent, pais_ori, import, neto,             porcen, indice, completo)"
            var_cadena = var_cadena + " values        ('" + Me.txt_pedido + "', '" + rs!numbodegarecibe + "','" + rs!numbodegadestino + "','" + rs!modeloproveedor + "','" + rs!numcodigocoppel + "','" + rs!numtallacoppel + "','" + rs!cantidadpedida + "', '" + rs!cantidadsurtida + "','" + rs!preciocosto + "', '" + rs!precioventa + "','0','0',     '',      " + rs!NUMPEDIDO + ",'0',     '0',        '0', '" + rs!numproveedor + "','',     '',       '',       '',       '',       '',       '',       '',       '',       '',        '',        '',        '',        '',    '',      '',     '',      '',        '',      '',        '','" + rs!neto + "','',     '0',  '') "
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
      Wend
      rs.Close
      rs.Open "select * from tb_pedido_original_coppel where archivo = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_fecha_inicio.SetFocus
      Else
         MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         Me.txt_pedido = ""
      End If
      rs.Close
   End If
End Sub


