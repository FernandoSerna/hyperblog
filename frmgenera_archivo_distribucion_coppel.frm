VERSION 5.00
Begin VB.Form frmgenera_archivo_distribucion_coppel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genera archivo de distribucion de COPPEL"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_generar_archivo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmgenera_archivo_distribucion_coppel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generar Archivo"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3165
      Picture         =   "frmgenera_archivo_distribucion_coppel.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3510
   End
   Begin VB.Frame Frame1 
      Caption         =   " Pedido "
      Height          =   1575
      Left            =   135
      TabIndex        =   0
      Top             =   435
      Width           =   3405
      Begin VB.TextBox txt_fecha_fin 
         Height          =   350
         Left            =   1665
         TabIndex        =   8
         Top             =   1065
         Width           =   1260
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   350
         Left            =   1665
         TabIndex        =   7
         Top             =   690
         Width           =   1260
      End
      Begin VB.TextBox txt_pedido 
         Height          =   350
         Left            =   1665
         TabIndex        =   2
         Top             =   315
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha fin:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicio:"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido de COPPEL:"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   390
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmgenera_archivo_distribucion_coppel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
         
         rs.Open "select * from tb_pedido_original_coppel where numpedido = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "DELETE FROM TB_aRCHIVO_DISTRIBUCION_COPPEL", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "SELECT DISTINCT INTE_CAR_NUMERO, INTE_ORS_ORDEN_SURTIDO, VCHA_sER_SERIE_ID, INTE_EMB_EMBARQUE, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_IMPORTE_IVA, destino FROM VW_ARCHIVO_DISTRIBUCION_COPPEL WHERE NUMPEDIDO = '" + Me.txt_pedido + "' and DTIM_CAR_FECHA >= " + var_fecha_inicio + " and dtim_Car_fecha <= " + var_fecha_fin + "-0.000001", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_orden_surtido = rsaux!INTE_ORS_ORDEN_SURTIDO
                  VAR_EMBARQUE = rsaux!INTE_EMB_EMBARQUE
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
                  VAR_ET_DATO6 = "FACTURA:" + CStr(rsaux!inte_car_numero)
                  rsaux2.Open "SELECT SUM(FLOA_SAL_CANTIDAD) FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO = " + CStr(rsaux!inte_car_numero) + " AND VCHA_SER_sERIE_ID = '" + rsaux!vcha_ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_cantidad = rsaux2(0).Value
                  VAR_ET_DATO7 = "UNIDS.FACTURA:" + CStr(var_cantidad)
                  rsaux2.Close
                  cnn.CommandTimeout = 360
                  rsaux2.Open "SELECT * FROM VW_ARCHIVO_DISTRIBUCION_COPPEL WHERE NUMPEDIDO = '" + Me.txt_pedido + "' and destino = '" + rsaux!Destino + "' and DTIM_CAR_FECHA >= " + var_fecha_inicio + " and dtim_Car_fecha <= " + var_fecha_fin + "-0.000001", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux2.EOF
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
                        rsaux4.Open "SELECT SUM(FLOA_SAL_CANTIDAD) FROM TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO = " + CStr(rsaux!inte_car_numero) + " AND VCHA_SER_sERIE_ID = '" + rsaux!vcha_ser_Serie_id + "' AND VCHA_ART_ARTICULO_ID = '" + rsaux2!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_cantidad_surtida = rsaux2!Cantidad
                        rsaux4.Close
                        var_cadena = "INSERT INTO TB_ARCHIVO_DISTRIBUCION_COPPEL (ARCHIVO, BODRECIBO, DESTINO, MODELOPROV, CODICOPPEL, TALLA_COPP, CANTPED, CANTSUR, COSTO, VENTA, LOTE, TOTLOTES, NUMFACTURA, NUMPEDIDO, IMPTEFACTU, IVAFACTURA, UNIDSFACTU, PROVEEDOR, ET_DATO1, ET_DATO2,"
                        var_cadena = var_cadena + "ET_DATO3, ET_DATO4, ET_DATO5, ET_DATO6, ET_DATO7, ET_DATO8, ET_DATO9, ET_DATO10, ET_DATO11, ET_DATO12, CDDESTINO, MARCA, FAMILIA, TRANSF, EMPAQUE, PEDIMENTO, PTO_ENT, PAIS_ORI, IMPORT, NETO, PORCEN, INDICE, COMPLETO, FECHAFAC)"
                        var_cadena = var_cadena + " Values ('" + rsaux2!archivo + "', '" + rsaux2!BODRECIBO + "','" + rsaux2!Destino + "', '" + rsaux2!MODELOPROV + "', '" + rsaux2!codicoppel + "', '" + rsaux2!TALLA_COPP + "', " + CStr(var_cantidad_surtida) + "," + CStr(var_cantidad_surtida) + "," + CStr(rsaux2!Costo) + ", " + CStr(rsaux2!VENTA) + ", '" + var_caja + "', '" + VAR_MAXIMO_cAJA + "', " + CStr(rsaux!inte_car_numero) + ", '" + rsaux2!numpedido + "', " + CStr(rsaux!floa_car_importe_neto) + ", " + CStr(rsaux!floa_car_importe_iva) + ", " + CStr(var_cantidad) + ", '" + rsaux2!proveedor + "',"
                        var_cadena = var_cadena + "'" + VAR_ET_DATO1 + "', '" + VAR_ET_DATO2 + "', '" + VAR_ET_DATO3 + "', '" + VAR_ET_DATO4 + "', '" + VAR_ET_DATO5 + "', '" + VAR_ET_DATO6 + "', '" + VAR_ET_DATO7 + "', '" + VAR_ET_DATO8 + "', '" + VAR_ET_DATO9 + "', '" + VAR_ET_DATO10 + "', '" + VAR_ET_DATO11 + "', '" + VAR_ET_DATO12 + "', '" + rsaux2!CDDESTINO + "', '" + rsaux2!marca + "', '" + rsaux2!familia + "', '" + rsaux2!TRANSF + "', '" + rsaux2!EMPAQUE + "', '" + rsaux2!PEDIMENTO + "', '" + rsaux2!PTO_ENT + "', '" + rsaux2!PAIS_ORI + "', '" + rsaux2!Import + "', " + CStr(rsaux2!NETO) + ", "
                        var_cadena = var_cadena + CStr(rsaux2!PORCEN) + ", " + CStr(rsaux2!indice) + ", '" + rsaux2!COMPLETO + "', NULL)"
                        rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.MoveNext
                  Wend
                  rsaux.MoveNext
            Wend
            rsaux.Close
            rsaux.Open "SELECT * FROM TB_PEDIDO_ORIGINAL_COPPEL WHERE NUMPEDIDO = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_Archivo = rsaux!archivo
                  var_pedido = rsaux!numpedido
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "SELECT * FROM TB_ARCHIVO_DISTRIBUCION_COPPEL WHERE ARCHIVO = '" + rsaux!archivo + "' AND BODRECIBO = '" + rsaux!BODRECIBO + "' AND DESTINO = '" + rsaux!Destino + "' AND MODELOPROV = '" + rsaux!MODELOPROV + "'AND CODICOPPEL = '" + rsaux!codicoppel + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux2.EOF Then
                     var_cadena = "INSERT INTO TB_ARCHIVO_DISTRIBUCION_COPPEL (ARCHIVO, BODRECIBO, DESTINO, MODELOPROV, CODICOPPEL, TALLA_COPP, CANTPED, CANTSUR, COSTO, VENTA, LOTE, TOTLOTES, NUMFACTURA, NUMPEDIDO, IMPTEFACTU, IVAFACTURA, UNIDSFACTU, PROVEEDOR, ET_DATO1, ET_DATO2,"
                     var_cadena = var_cadena + "ET_DATO3, ET_DATO4, ET_DATO5, ET_DATO6, ET_DATO7, ET_DATO8, ET_DATO9, ET_DATO10, ET_DATO11, ET_DATO12, CDDESTINO, MARCA, FAMILIA, TRANSF, EMPAQUE, PEDIMENTO, PTO_ENT, PAIS_ORI, IMPORT, NETO, PORCEN, INDICE, COMPLETO, FECHAFAC)"
                     var_cadena = var_cadena + " Values ('" + rsaux!archivo + "', '" + rsaux!BODRECIBO + "','" + rsaux!Destino + "', '" + rsaux!MODELOPROV + "', '" + rsaux!codicoppel + "', '" + rsaux!TALLA_COPP + "', " + CStr(rsaux!cantped) + ",0," + CStr(rsaux!Costo) + ", " + CStr(rsaux!VENTA) + ", '', '', '', '" + rsaux!numpedido + "', 0, 0, 0, '" + rsaux!proveedor + "',"
                     var_cadena = var_cadena + "'" + rsaux!ET_DATO1 + "', '" + rsaux!ET_DATO2 + "', '" + rsaux!ET_DATO3 + "', '" + rsaux!ET_DATO4 + "', '" + rsaux!ET_DATO5 + "', '" + rsaux!ET_DATO6 + "', '" + rsaux!ET_DATO7 + "', '" + rsaux!ET_DATO8 + "', '" + rsaux!ET_DATO9 + "', '" + rsaux!ET_DATO10 + "', '" + rsaux!ET_DATO11 + "', '" + rsaux!ET_DATO12 + "', '" + rsaux!CDDESTINO + "', '" + rsaux!marca + "', '" + rsaux!familia + "', '" + rsaux!TRANSF + "', '" + rsaux!EMPAQUE + "', '" + rsaux!PEDIMENTO + "', '" + rsaux!PTO_ENT + "', '" + rsaux!PAIS_ORI + "', '" + rsaux!Import + "', " + CStr(rsaux!NETO) + ", "
                     var_cadena = var_cadena + CStr(rsaux!PORCEN) + ", " + CStr(rsaux!indice) + ", '" + rsaux!COMPLETO + "', NULL)"
                     rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rsaux.MoveNext
            Wend
            rsaux.Close
            var_eliminar = DeleteFile("c:\coppel\archivos\tem_" + Trim(var_Archivo) + ".dbf")
            var_eliminar = DeleteFile("c:\coppel\archivos\" + Trim(var_Archivo) + ".dbf")
            var_eliminar = DeleteFile("c:\coppel\archivos\" + Trim(var_Archivo) + ".dbf")
            var_copia = CopyFile("c:\coppel\archivos\temp_archivo.dbf", "c:\coppel\archivos\tem_" + Trim(var_Archivo) + ".dbf", 1)
            var_copia = CopyFile("c:\coppel\archivos\temp_archivo.dbf", "c:\coppel\archivos\tem_" + Trim(var_Archivo) + ".dbf", 1)
            rsaux5.Open "DELETE FROM tem_" + var_Archivo, var_tabla, adOpenDynamic, adLockOptimistic
            rsaux.Open "select * from tb_archivo_distribucion_coppel WHERE NUMPEDIDO = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_cadena = "INSERT INTO tem_" + var_Archivo + ".dbf (BODRECIBO, DESTINO, MODELOPROV, CODICOPPEL, TALLA_COPP, CANTPED, CANTSUR, COSTO, VENTA, LOTE, TOTLOTES, NUMFACTURA, NUMPEDIDO, IMPTEFACTU, IVAFACTURA, UNIDSFACTU, PROVEEDOR, ET_DATO1, ET_DATO2, ET_DATO3, ET_DATO4, ET_DATO5, ET_DATO6, ET_DATO7, ET_DATO8, ET_DATO9, ET_DATO10, ET_DATO11, ET_DATO12, CDDESTINO, MARCA, FAMILIA, TRANSF, EMPAQUE, PEDIMENTO, PTO_ENT, PAIS_ORI, IMPORT, NETO, PORCEN, INDICE, COMPLETO, fechafac)"
                  var_cadena = var_cadena + " Values ( '" + rsaux!BODRECIBO + "','" + rsaux!Destino + "', '" + rsaux!MODELOPROV + "', '" + rsaux!codicoppel + "', '" + rsaux!TALLA_COPP + "', " + CStr(rsaux!cantped) + "," + CStr(rsaux!CANTSUR) + ", " + CStr(rsaux!Costo) + ", " + CStr(rsaux!VENTA) + ", '" + rsaux!lote + "', '" + rsaux!TOTLOTES + "', '" + rsaux!numfactura + "',   '" + rsaux!numpedido + "'," + CStr(rsaux!IMPTEFACTU) + ", " + CStr(rsaux!IVAFACTURA) + "," + CStr(rsaux!UNIDSFACTU) + ","
                  var_cadena = var_cadena + "'" + rsaux!proveedor + "', '" + rsaux!ET_DATO1 + "','" + rsaux!ET_DATO2 + "', '" + rsaux!ET_DATO3 + "', '" + rsaux!ET_DATO4 + "', '" + rsaux!ET_DATO5 + "', '" + rsaux!ET_DATO6 + "', '" + rsaux!ET_DATO7 + "', '" + rsaux!ET_DATO8 + "', '" + rsaux!ET_DATO9 + "', '" + rsaux!ET_DATO10 + "', '" + rsaux!ET_DATO11 + "', '" + rsaux!ET_DATO12 + "', '" + rsaux!CDDESTINO + "', '" + rsaux!marca + "', '" + rsaux!familia + "', '" + rsaux!TRANSF + "', '" + rsaux!EMPAQUE + "', '" + rsaux!PEDIMENTO + "', '" + rsaux!PTO_ENT + "', '" + rsaux!PAIS_ORI + "', '" + rsaux!Import + "',"
                  var_cadena = var_cadena + CStr(rsaux!NETO) + ", " + CStr(rsaux!PORCEN) + ", " + CStr(rsaux!indice) + ", '" + rsaux!COMPLETO + "', ctod(" + """" + " / / " + """" + "))"
                  rsaux4.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                  rsaux.MoveNext
            Wend
            var_copia = CopyFile("c:\coppel\archivos\tem_" + var_Archivo + ".dbf", "c:\coppel\archivos\" + Trim(var_Archivo) + ".dbf", 1)
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select distinct destino, totlotes, numfactura, imptefactu, ivafactura, unidsfactu from c:\coppel\archivos\" + Trim(var_Archivo) + " where allt(totlotes)<> ''", var_tabla, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  rsaux2.Open "update " + Trim(var_Archivo) + " set totlotes = '" + rsaux!TOTLOTES + "', numfactura = '" + rsaux!numfactura + "', imptefactu = " + CStr(rsaux!IMPTEFACTU) + ", ivafactura = " + CStr(rsaux!IVAFACTURA) + ", unidsfactu = " + CStr(rsaux!UNIDSFACTU) + " where destino = '" + rsaux!Destino + "' and allt(numfactura) = '' ", var_tabla, adOpenDynamic, adLockOptimistic
                  rsaux.MoveNext
            Wend
            rsaux.Close
            rsaux.Open "select * from " + Trim(var_Archivo) + " order by destino, lote into table p1" + Trim(var_pedido), var_tabla, adOpenDynamic, adLockOptimistic
         
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
   Call activa_forma(var_activa_forma_existencias_generales)
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
      rs.Open "select * from tb_pedido_original_coppel where numpedido = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_fecha_inicio.SetFocus
      Else
         MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         Me.txt_pedido = ""
      End If
      rs.Close
   End If
End Sub

