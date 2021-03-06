VERSION 5.00
Begin VB.Form frmreporte_entradas_traspasos_compras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repoerte de entradas por traspaso y compras"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "EI"
      Height          =   315
      Left            =   1755
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Entradas intercompaņias"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "DP"
      Height          =   315
      Left            =   1425
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Devoluciones a proveedores"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "TE"
      Height          =   315
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Entradas para traspaso"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "TR"
      Height          =   315
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salidas de traspasos"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "C"
      Height          =   315
      Left            =   435
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Entradas por compra"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   120
      TabIndex        =   2
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
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
      Left            =   4005
      Picture         =   "frmreporte_entradas_traspasos_compras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Caption         =   "T"
      Height          =   315
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Reporte de entradas por traspasos de planta"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_entradas_traspasos_compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim aņo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_aņo = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR) "
            var_cadena = var_cadena + "SELECT  " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.0000001,dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO , dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, '','', dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_emo_referencia, dbo.tb_articulos.MONE_ART_PRECIO_base FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN"
            var_cadena = var_cadena + " dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN  dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'ETMP') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + "-.00001) ORDER BY DTIM_EMO_fECHA"
            Text1 = "SELECT  " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.0000001,dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO , dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, '','', dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_emo_referencia, dbo.tb_articulos.MONE_ART_PRECIO_base FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN"
            Text1 = Text1 + " dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN  dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'ETMP') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + "-.00001) ORDER BY DTIM_EMO_fECHA"
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_traspasos.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte entradas por traspasos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("ŋDesea exportar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_traspasos.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_entradas_traspasos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim aņo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_aņo = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR) "
            
            var_cadena = var_cadena + " SELECT     TOP 100 PERCENT " + CStr(var_consecutivo) + " AS Expr1, " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_PROVEEDORES.VCHA_PRO_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_FACTURA, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE "
            var_cadena = var_cadena + " FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN  dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_PROVEEDORES ON "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_PROVEEDORES.VCHA_PRO_PROVEEDOR_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'EC') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            
            
            var_cadena2 = " SELECT     TOP 100 PERCENT " + CStr(var_consecutivo) + " AS Expr1, " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_PROVEEDORES.VCHA_PRO_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_FACTURA, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE "
            var_cadena2 = var_cadena2 + " FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN  dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_PROVEEDORES ON "
            var_cadena2 = var_cadena2 + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_PROVEEDORES.VCHA_PRO_PROVEEDOR_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'EC') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            Text1 = var_cadena2
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_compras.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte entradas por compra"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("ŋDesea exportar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_compras.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_entradas_compras_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim aņo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_aņo = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            If var_empresa = "18" Then
               var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR, VCHA_TEM_NOMBRE_ALMACEN_ORIGEN) "
               var_cadena = var_cadena + " SELECT     " + CStr(var_consecutivo) + " , " + var_fecha_inicio + ", " + var_fecha_fin + " - .0000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_ALMACEN_DESTINO, '' AS Expr5, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, TB_ALMACENES_1.VCHA_ALM_NOMBRE AS ALMACEN_DESTINO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON "
               var_cadena = var_cadena + "  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES TB_ALMACENES_1 ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_ALMACEN_DESTINO = TB_ALMACENES_1.VCHA_ALM_ALMACEN_ID "
               var_cadena = var_cadena + " WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'TR') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND  (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            End If
            If var_empresa = "06" Then
               var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR, VCHA_TEM_NOMBRE_ALMACEN_ORIGEN) "
               var_cadena = var_cadena + " SELECT     " + CStr(var_consecutivo) + " , " + var_fecha_inicio + ", " + var_fecha_fin + " - .0000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, '' AS Expr5, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, TB_ALMACENES_1.VCHA_ALM_NOMBRE AS ALMACEN_DESTINO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON "
               var_cadena = var_cadena + "  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES TB_ALMACENES_1 ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_ALMACEN_DESTINO = TB_ALMACENES_1.VCHA_ALM_ALMACEN_ID "
               var_cadena = var_cadena + " WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'DPL') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND  (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            End If
            
            
            
            
            Text1 = var_cadena2
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_traspasos_SALIDA.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de salidas"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("ŋDesea exportar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_traspasos_SALIDA.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_entradas_traspasos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command3_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim aņo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_aņo = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            If var_empresa = "18" Then
               var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR, VCHA_TEM_NOMBRE_ALMACEN_ORIGEN) "
               var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + " - .0000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_ENTRADAS.VCHA_ENT_ALMACEN_ORIGEN,  dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN AS Expr4, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_MOVIMIENTO_ORIGEN AS Expr5, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, TB_ALMACENES_1.VCHA_ALM_NOMBRE AS ALMACEN_ORIGEN "
               var_cadena = var_cadena + " FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES TB_ALMACENES_1 ON "
               var_cadena = var_cadena + " dbo.TB_ENTRADAS.VCHA_ENT_ALMACEN_ORIGEN = TB_ALMACENES_1.VCHA_ALM_ALMACEN_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'TE') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA             "
            End If
            If var_empresa = "06" Then
               var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR, VCHA_TEM_NOMBRE_ALMACEN_ORIGEN) "
               
               
               var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + " - .0000001,dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID,dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_ENTRADAS.VCHA_ENT_ALMACEN_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN AS Expr4, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_MOVIMIENTO_ORIGEN AS Expr5, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA AS ALMACEN_ORIGEN_2 FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN"
               var_cadena = var_cadena + " dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'ETP') AND "
               var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") And (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - 0.00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
               
               
               'var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + " - .0000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_ENTRADAS.VCHA_ENT_ALMACEN_ORIGEN,  dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN AS Expr4, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_MOVIMIENTO_ORIGEN AS Expr5, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, TB_ALMACENES_1.VCHA_ALM_NOMBRE AS ALMACEN_ORIGEN "
               'var_cadena = var_cadena + " FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES TB_ALMACENES_1 ON "
               'var_cadena = var_cadena + " dbo.TB_ENTRADAS.VCHA_ENT_ALMACEN_ORIGEN = TB_ALMACENES_1.VCHA_ALM_ALMACEN_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'ETP') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA             "
            End If
            
            Text1 = var_cadena
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_traspasos_ENTRADA.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte entradas por traspasos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("ŋDesea exportar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_traspasos_ENTRADA.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_entradas_traspasos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command4_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim aņo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_aņo = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR, VCHA_TEM_NOMBRE_ALMACEN_ORIGEN) "
            
            
            var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + " - .0000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID,vcha_Alm_nombre, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, dbo.TB_SALIDAS.FLOA_SAL_COSTO, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_PROVEEDORES.VCHA_PRO_PROVEEDOR_ID, dbo.TB_PROVEEDORES.VCHA_PRO_NOMBRE , dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_REFERENCIA, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE, '' FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN "
            var_cadena = var_cadena + " dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_PROVEEDORES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_PROVEEDORES.VCHA_PRO_PROVEEDOR_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'DP') AND "
            var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA "
            
            Text1 = var_cadena
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_proveedores.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte devoluciones a proveedores"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("ŋDesea exportar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_proveedores.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_devoliciones_proveedores_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command5_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim aņo As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_aņo = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            var_cadena = " insert into TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS (INTE_TEM_CONSECUTIVO, DTIM_tEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO , FLOA_ENT_PRECIO, VCHA_TEM_ORIGEN, VCHA_TEM_NUMERO_ORIGEN, VCHA_TEM_REFERENCIA_EXTRA, MONE_ART_PRECIO_ESTANDAR) "
            
            'var_cadena = var_cadena + " SELECT     TOP 100 PERCENT " + CStr(var_consecutivo) + " AS Expr1, " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_PROVEEDORES.VCHA_PRO_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_FACTURA, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE "
            'var_cadena = var_cadena + " FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN  dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_PROVEEDORES ON "
            'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_PROVEEDORES.VCHA_PRO_PROVEEDOR_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            
            
            var_cadena = var_cadena + "sELECT     TOP 100 PERCENT " + CStr(var_consecutivo) + " AS Expr1, " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE,  dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID, dbo.tb_ARTICULOS.VCHA_ART_NOMBRE_ESPAŅOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD, dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMO_FACTURA, dbo.tb_ARTICULOS.MONE_ART_PRECIO_BASE "
            var_cadena = var_cadena + " FROM         dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.tb_ARTICULOS ON dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID = dbo.tb_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES ON "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID WHERE     (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + " - .00001) ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_aņo = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_aņo + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_intercompaņias.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte entradas por compra"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("ŋDesea exportar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos_compra_intercompaņias.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_entradas_compras_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_TRASPASOS_COMPRAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub



