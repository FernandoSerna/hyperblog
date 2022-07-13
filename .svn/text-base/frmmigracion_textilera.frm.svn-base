VERSION 5.00
Begin VB.Form frmmigracion_textilera 
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "10.- Inventario Inicial Salidas"
      Height          =   720
      Left            =   4110
      TabIndex        =   9
      Top             =   1005
      Width           =   3405
   End
   Begin VB.CommandButton cmd_inventario_inicial 
      Caption         =   "9.- Inventario Inicial"
      Height          =   720
      Left            =   4125
      TabIndex        =   8
      Top             =   150
      Width           =   3405
   End
   Begin VB.CommandButton COM_LISTA_PRECIOS 
      Caption         =   "8.- Lista de Precios"
      Height          =   720
      Left            =   585
      TabIndex        =   7
      Top             =   6045
      Width           =   3405
   End
   Begin VB.CommandButton cdm_migrar_promociones 
      Caption         =   "7.- promociones"
      Height          =   720
      Left            =   585
      TabIndex        =   6
      Top             =   5235
      Width           =   3405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6.- equivalencias Reclasificacion 7"
      Height          =   720
      Left            =   585
      TabIndex        =   5
      Top             =   4410
      Width           =   3405
   End
   Begin VB.CommandButton cmd_codigos_reclasificacion 
      Caption         =   "5.- equivalencias Reclasificacion"
      Height          =   720
      Left            =   585
      TabIndex        =   4
      Top             =   3600
      Width           =   3405
   End
   Begin VB.CommandButton cmd_reclasificacion 
      Caption         =   "4.-  equivalencias Reclasificacion"
      Height          =   720
      Left            =   585
      TabIndex        =   3
      Top             =   2790
      Width           =   3405
   End
   Begin VB.CommandButton cmd_equivalencias 
      Caption         =   "3.-  migracion equivalencias"
      Height          =   720
      Left            =   585
      TabIndex        =   2
      Top             =   1980
      Width           =   3405
   End
   Begin VB.CommandButton cmd_migracion_articulos 
      Caption         =   "2.-   Migración de Artículos"
      Height          =   720
      Left            =   585
      TabIndex        =   1
      Top             =   1035
      Width           =   3405
   End
   Begin VB.CommandButton cmd_migrar_articulos 
      Caption         =   "1.-   migracion conceptos"
      Height          =   720
      Left            =   585
      TabIndex        =   0
      Top             =   150
      Width           =   3405
   End
   Begin VB.Label Label2 
      Height          =   690
      Left            =   4125
      TabIndex        =   11
      Top             =   2790
      Width           =   1395
   End
   Begin VB.Label Label1 
      Height          =   480
      Left            =   4155
      TabIndex        =   10
      Top             =   2040
      Width           =   1290
   End
End
Attribute VB_Name = "frmmigracion_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection

Private Sub cdm_migrar_promociones_Click()
   Dim verificador As Integer
    rs.Open "select distinct agente, codigo, descuento0, descuento1, descuento2, descuento3, descuento4, descuento5, descuento6, descuento7, descuento8, descuento9, tipoclient, vigencia_i, vigencia_f from promagen where vigencia_f >= date() and tipoclient <> 'A'", var_tabla, adOpenDynamic, adLockOptimistic
    Label1.Caption = CStr(rs.RecordCount)
    Me.Refresh
    var_contador = 0
    While Not rs.EOF
          var_contador = var_contador + 1
          Label2.Caption = CStr(var_contador)
          Me.Refresh
          If Trim(rs!agente) = "2" Then
             canal = "29"
          End If
          If Trim(rs!agente) = "1" Then
             canal = "27"
          End If
          If Trim(rs!agente) = "3" Then
             canal = "23"
          End If
          If Trim(rs!agente) = "4" Then
             canal = "26"
          End If
          If Trim(rs!agente) = "6" Then
             canal = "24"
          End If
          If Trim(rs!agente) = "7" Then
             canal = "30"
          End If
          If Trim(rs!agente) = "154" Then
             canal = "25"
          End If
          If Trim(rs!agente) = "265" Then
             canal = "28"
          End If
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "0"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento0) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento0) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "1"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento1) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento1) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
           
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "2"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento2) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento2) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
            
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "3"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento3) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento3) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
            
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "4"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento4) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento4) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "5"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento5) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento5) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "6"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento6) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento6) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "7"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento7) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento7) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "8"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento8) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento8) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          var_codigo = rs!codigo
          var_codigo = Trim(var_codigo) + "9"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          var_codigo = var_codigo + Trim(CStr(verificador))
          'rsaux5.Open "select * from tb_descuentos_promociones where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'If Not rsaux5.EOF Then
          '   rsaux.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = '" + CStr(rs!vigencia_i) + "', dtim_dpr_fecha_fin = '" + CStr(rs!vigencia_f) + "', floa_dpr_descuento = " + CStr(rs!descuento9) + " where vcha_can_canal_venta_id = '" + canal + "' and  vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
          'Else
             rsaux.Open "insert tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_Art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + canal + "','" + var_codigo + "','" + CStr(rs!vigencia_i) + "','" + CStr(rs!vigencia_f) + "'," + CStr(rs!descuento9) + ")", cnn, adOpenDynamic, adLockOptimistic
          'End If
          'rsaux5.Close
          
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cmd_codigos_reclasificacion_Click()
   Dim verificador As Integer
   rs.Open "select codigo, codigorecl from articulos where len(codigorecl) > 0", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "0"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            var_codigo = rs!codigorecl
            var_codigo = Trim(var_codigo) + "0"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_2 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "1"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "2"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "3"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "4"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "5"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "6"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "6"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "7"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "8"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            var_codigo = rs!codigo
            var_codigo = Trim(var_codigo) + "9"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            var_codigo_1 = var_codigo
            
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RETEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "insert into tb_Reclasificacion (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_rec_codigo_general) values ('RVTEX', '" + var_codigo_1 + "', '" + var_codigo_2 + "')", cnn, adOpenDynamic, adLockOptimistic
            
            
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_equivalencias_Click()
   Dim verificador As Integer
   rs.Open "select * from equivalencias where len(allt(codigo2)) = 10", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select * from tb_equivalencias where vcha_art_articulo_id = '" + Trim(rs!codigo2) + "' and vcha_equ_codigo_equivalente = '" + rs!codigo1 + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            var_t = IIf(IsNull(rs!tipo), 0, rs!tipo)
            var_codigo = rs!codigo2
            var_codigo = Trim(var_codigo) + "0"
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
            verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If verificador = 10 Then
               verificador = 0
            End If
          
            var_codigo = var_codigo + Trim(CStr(verificador))
            
            rsaux2.Open "insert into tb_equivalencias (vcha_equ_codigo_equivalente, vcha_art_Articulo_id, inte_equ_codigo_interno ,inte_equ_codigo_bulto) values ('" + rs!codigo1 + "', '" + var_codigo + "', 0," + CStr(var_t) + ")", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_inventario_inicial_Click()
Dim var_año As Integer
Dim var_almacen_Destino As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_numero_causa As Integer
Dim var_elimina As Boolean
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_renglon As Double
Dim verificador As Integer
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   var_si = 6
   If var_si = 6 Then
   bandera_suma = False
   var_inserta = False
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "PTTEX"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select existen.codigo, cantidad, a.costo, preciou from existen, articulos a where cve_almace = 4 and cantidad > 0 and subs(existen.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "RETEX"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select existen.codigo, cantidad, a.costo, preciou from existen, articulos a where cve_almace = 1 and cantidad > 0 and subs(existen.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "RVTEX"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select existen.codigo, cantidad, a.costo, preciou from existen, articulos a where cve_almace = 5 and cantidad > 0 and subs(existen.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   End If
   
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00095"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '01' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   
   
   
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00096"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '10' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   
   
   
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00099"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '12' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00099"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '14' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
   
   
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00098"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '13' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close

   
      var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00097"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '15' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close

   
   
   MsgBox "Se a terminado de subir los movimientos", vbOKOnly, "ATENCION"
 End Sub

Private Sub cmd_migracion_articulos_Click()
    Dim verificador As Integer
    rs.Open "select cve_linea, descripcio from lineas", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "select * from tb_lineas where vcha_lin_linea_id = '" + CStr(rs!cve_linea) + "'", cnn, adOpenDynamic, adLockOptimistic
          If rsaux.EOF Then
             rsaux2.Open "insert into tb_lineas (vcha_lin_linea_id, vcha_lin_nombre) values ('" + CStr(rs!cve_linea) + "','" + rs!descripcio + "')", cnn, adOpenDynamic, adLockOptimistic
          End If
          rsaux.Close
          rs.MoveNext
    Wend
    rs.Close
    rs.Open "select cve_catalo, descripcio from catalogos", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "select * from tb_catalogos where vcha_cat_catalogo_id = '" + rs!cve_catalo + "'", cnn, adOpenDynamic, adLockOptimistic
          If rsaux.EOF Then
             rsaux2.Open "insert into tb_catalogos (vcha_Cat_catalogo_id, vcha_cat_nombre) values ('" + rs!cve_catalo + "','" + rs!descripcio + "')", cnn, adOpenDynamic, adLockOptimistic
          End If
          rsaux.Close
          rs.MoveNext
    Wend
    rs.Close
    var_i = 0
    var_j = 0
    Dim x As Integer
    
    rs.Open "select distinct codigo from existen where val(subs(codigo,1,1)) <> 6 and val(subs(codigo,1,1)) <> 7 and val(subs(codigo,1,1)) <> 5 AND LEN(ALLT(CODIGO))>0", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          'For x = 0 To 9
             var_i = var_i + 1
             var_codigo = rs!codigo
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
             verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
             If verificador = 10 Then
                verificador = 0
             End If
          
             var_codigo = var_codigo + Trim(CStr(verificador))
          
             var_descuento = CInt(Mid(Trim(var_codigo), 11, 1)) * 10
          
             var_codigo_2 = Left(rs!codigo, 10)
          
          
             var_tipo = Val(Left(Trim(var_codigo_2), 1))
             VAR_DIVISION = Val(Mid(Trim(var_codigo_2), 2, 2))
             var_subdivision = Val(Mid(var_codigo_2, 4, 2))
             VAR_ESTAMPADO = Val(Mid(var_codigo_2, 6, 5))
           
             zzzz = 0
             rsaux.Open "select * from articulos where tipo = " + CStr(var_tipo) + " and division = " + CStr(VAR_DIVISION) + " and subdivisio = " + CStr(var_subdivision) + " and estampado = " + CStr(VAR_ESTAMPADO), cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux.EOF Then
                var_costo = rsaux!costo_prov
                var_precio = rsaux!preciou
                var_precio = var_precio * ((100 - var_descuento) / 100)
                var_tipo = CStr(rsaux!tipo)
                If Len(CStr(rsaux!division)) = 1 Then
                   VAR_DIVISION = "0" + Trim(CStr(rsaux!division))
                Else
                   VAR_DIVISION = Trim(CStr(rsaux!division))
                End If
                If Len(CStr(rsaux!subdivisio)) = 1 Then
                   var_subdivision = "0" + Trim(CStr(rsaux!subdivisio))
                Else
                   var_subdivision = Trim(CStr(rsaux!subdivisio))
                End If
                If Len(CStr(rsaux!estampado)) = 1 Then
                   VAR_ESTAMPADO = "0000" + Trim(CStr(rsaux!estampado))
                Else
                   If Len(CStr(rsaux!estampado)) = 2 Then
                      VAR_ESTAMPADO = "000" + Trim(CStr(rsaux!estampado))
                   Else
                      If Len(CStr(rsaux!estampado)) = 3 Then
                         VAR_ESTAMPADO = "00" + Trim(CStr(rsaux!estampado))
                      Else
                        If Len(CStr(rsaux!estampado)) = 4 Then
                           VAR_ESTAMPADO = "0" + Trim(CStr(rsaux!estampado))
                        Else
                           If Len(CStr(rsaux!estampado)) = 5 Then
                              VAR_ESTAMPADO = Trim(CStr(rsaux!estampado))
                           End If
                        End If
                      End If
                   End If
                End If
                var_linea = rsaux!cve_linea
                var_catalogo = rsaux!cve_catalo
                
                'rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                'If rsaux2.EOF Then
                   var_j = var_j + 1
             
                   rsaux4.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + var_tipo + "' and vcha_div_division_id = '" + VAR_DIVISION + "' and vcha_sub_subdivision_id = '" + var_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      var_desc_sub = IIf(IsNull(rsaux4!VCHA_SUB_NOMBRE), "", rsaux4!VCHA_SUB_NOMBRE)
                   Else
                      var_desc_sub = ""
                   End If
                   rsaux4.Close
                
                   rsaux4.Open "select * from tb_estampados where vcha_est_estampado_id = '" + VAR_ESTAMPADO + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      var_desc_est = IIf(IsNull(rsaux4!VCHA_EST_NOMBRE), "", rsaux4!VCHA_EST_NOMBRE)
                   Else
                      var_desc_est = ""
                   End If
                   rsaux4.Close
                   VAR_DESCRIPCION = Trim(var_desc_sub) + " " + Trim(var_desc_est)
                   bulto = rsaux!bulto
                   If rsaux3.State = 1 Then
                      rsaux3.Close
                   End If
                   rsaux3.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                   If rsaux3.EOF Then
                      var_cadena = "insert into tb_Articulos (vcha_art_articulo_id, vcha_art_nombre_español, mone_art_precio_base, mone_art_costo_estandar, vcha_art_catalogo_inicio, vcha_Art_catalogo_vigente, vcha_lin_linea_id, vcha_uni_unidad_id, inte_art_salida_masiva, vcha_Art_codigo_externo, vcha_tpr_tipo_producto_id, vcha_div_division_id, vcha_sub_subdivision_id, vcha_est_estampado_id) values "
                      var_cadena = var_cadena + "('" + var_codigo + "','" + Left(VAR_DESCRIPCION, 50) + "', " + CStr(var_precio) + "," + CStr(var_costo) + ",'" + var_catalogo + "','" + var_catalogo + "', '" + var_linea + "', ''," + CStr(bulto) + ",'','" + var_tipo + "','" + VAR_DIVISION + "','" + var_subdivision + "','" + VAR_ESTAMPADO + "')"
                      rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                   End If
                'End If
                'rsaux2.Close
             
             End If
             rsaux.Close
          'Next x
          rs.MoveNext
    Wend
    rs.Close
    
    
    x = 1
    If x = 0 Then
    
    rs.Open "select distinct codigo from articulos where tipo= 6 or tipo = 7 or tipo= 5", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          For x = 0 To 9
             var_i = var_i + 1
             var_codigo = rs!codigo + Trim(CStr(x))
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
             verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
             If verificador = 10 Then
                verificador = 0
             End If
          
             var_codigo = var_codigo + Trim(CStr(verificador))
          
             var_descuento = CInt(Mid(var_codigo, 11, 1)) * 10
          
             var_codigo_2 = Left(rs!codigo, 10)
          
          
             var_tipo = Val(Left(Trim(var_codigo_2), 1))
             VAR_DIVISION = Val(Mid(Trim(var_codigo_2), 2, 2))
             var_subdivision = Val(Mid(var_codigo_2, 4, 2))
             VAR_ESTAMPADO = Val(Mid(var_codigo_2, 6, 5))
           
             
             rsaux.Open "select * from articulos where tipo = " + CStr(var_tipo) + " and division = " + CStr(VAR_DIVISION) + " and subdivisio = " + CStr(var_subdivision) + " and estampado = " + CStr(VAR_ESTAMPADO), cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux.EOF Then
                var_costo = rsaux!costo_prov
                var_precio = rsaux!preciou
                var_precio = var_precio * ((100 - var_descuento) / 100)
                var_tipo = CStr(rsaux!tipo)
                If Len(CStr(rsaux!division)) = 1 Then
                   VAR_DIVISION = "0" + Trim(CStr(rsaux!division))
                Else
                   VAR_DIVISION = Trim(CStr(rsaux!division))
                End If
                If Len(CStr(rsaux!subdivisio)) = 1 Then
                   var_subdivision = "0" + Trim(CStr(rsaux!subdivisio))
                Else
                   var_subdivision = Trim(CStr(rsaux!subdivisio))
                End If
                If Len(CStr(rsaux!estampado)) = 1 Then
                   VAR_ESTAMPADO = "0000" + Trim(CStr(rsaux!estampado))
                Else
                   If Len(CStr(rsaux!estampado)) = 2 Then
                      VAR_ESTAMPADO = "000" + Trim(CStr(rsaux!estampado))
                   Else
                      If Len(CStr(rsaux!estampado)) = 3 Then
                         VAR_ESTAMPADO = "00" + Trim(CStr(rsaux!estampado))
                      Else
                        If Len(CStr(rsaux!estampado)) = 4 Then
                           VAR_ESTAMPADO = "0" + Trim(CStr(rsaux!estampado))
                        Else
                           If Len(CStr(rsaux!estampado)) = 5 Then
                              VAR_ESTAMPADO = Trim(CStr(rsaux!estampado))
                           End If
                        End If
                      End If
                   End If
                End If
                var_linea = rsaux!cve_linea
                var_catalogo = rsaux!cve_catalo
                
                'rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                'If rsaux2.EOF Then
                   var_j = var_j + 1
             
                   rsaux4.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + var_tipo + "' and vcha_div_division_id = '" + VAR_DIVISION + "' and vcha_sub_subdivision_id = '" + var_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      var_desc_sub = IIf(IsNull(rsaux4!VCHA_SUB_NOMBRE), "", rsaux4!VCHA_SUB_NOMBRE)
                   Else
                      var_desc_sub = ""
                   End If
                   rsaux4.Close
                
                   rsaux4.Open "select * from tb_estampados where vcha_est_estampado_id = '" + VAR_ESTAMPADO + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux4.EOF Then
                      var_desc_est = IIf(IsNull(rsaux4!VCHA_EST_NOMBRE), "", rsaux4!VCHA_EST_NOMBRE)
                   Else
                      var_desc_est = ""
                   End If
                   rsaux4.Close
                   VAR_DESCRIPCION = Trim(var_desc_sub) + " " + Trim(var_desc_est)
                   bulto = rsaux!bulto
                   
                   var_cadena = "insert into tb_Articulos (vcha_art_articulo_id, vcha_art_nombre_español, mone_art_precio_base, mone_art_costo_estandar, vcha_art_catalogo_inicio, vcha_Art_catalogo_vigente, vcha_lin_linea_id, vcha_uni_unidad_id, inte_art_salida_masiva, vcha_Art_codigo_externo, vcha_tpr_tipo_producto_id, vcha_div_division_id, vcha_sub_subdivision_id, vcha_est_estampado_id) values "
                   var_cadena = var_cadena + "('" + var_codigo + "','" + Left(VAR_DESCRIPCION, 50) + "', " + CStr(var_precio) + "," + CStr(var_costo) + ",'" + var_catalogo + "','" + var_catalogo + "', '" + var_linea + "', ''," + CStr(bulto) + ",'','" + var_tipo + "','" + VAR_DIVISION + "','" + var_subdivision + "','" + VAR_ESTAMPADO + "')"
                   rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                
                'End If
                'rsaux2.Close
             
             End If
             rsaux.Close
          Next x
          rs.MoveNext
    Wend
    rs.Close
    End If
    MsgBox CStr(var_i) + "      " + CStr(var_j)
End Sub

Private Sub cmd_migrar_articulos_Click()
   rs.Open "delete from tb_tipos_productos", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_divisiones ", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_subdivisiones ", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_estampados", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select * from tipoprod", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into tb_tipos_productos (vcha_tpr_tipo_producto_id, vcha_tpr_nombre) values ('" + CStr(rs!cve_produc) + "', '" + rs!descripcio + "')", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close

   rs.Open "select round(depende,0) as depende, round(cve_divisi,0) as cve_divisi, descripcio from division", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If Len(CStr(rs!cve_divisI)) = 1 Then
            VAR_DIVISION = "0" + CStr(rs!cve_divisI)
         Else
            VAR_DIVISION = CStr(rs!cve_divisI)
         End If
         rsaux.Open "insert into tb_divisiones (vcha_tpr_tipo_producto_id, vcha_div_division_id, vcha_div_nombre) values ('" + CStr(rs!depende) + "', '" + VAR_DIVISION + "', '" + CStr(rs!descripcio) + "')", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close

   rs.Open "select round(cve_divi,0) as cve_divi, round(depende,0) as depende, round(depende2,0) as depende_2,descripcio from subdivis", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If Len(CStr(rs!cve_divi)) = 1 Then
            var_subdivision = "0" + CStr(rs!cve_divi)
         Else
            var_subdivision = CStr(rs!cve_divi)
         End If
         
         If Len(CStr(rs!depende_2)) = 1 Then
            VAR_DIVISION = "0" + CStr(rs!depende_2)
         Else
            VAR_DIVISION = CStr(rs!depende_2)
         End If
         rsaux2.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + CStr(rs!depende) + "' and vcha_div_division_id = '" + VAR_DIVISION + "' and vcha_sub_subdivision_id = '" + var_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux2.EOF Then
            rsaux.Open "insert into tb_subdivisiones (vcha_tpr_tipo_producto_id, vcha_div_division_id, vcha_sub_subdivision_id, vcha_sub_nombre) values ('" + CStr(rs!depende) + "', '" + VAR_DIVISION + "', '" + var_subdivision + "', '" + CStr(rs!descripcio) + " ')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux2.Close
         rs.MoveNext
   Wend
   rs.Close

   rs.Open "select round(cve_divi,0) as cve_divi, round(depende,0) as depende, round(depende2,0) as depende_2,descripcio from subdivis", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If Len(CStr(rs!cve_divi)) = 1 Then
            var_subdivision = "0" + CStr(rs!cve_divi)
         Else
            var_subdivision = CStr(rs!cve_divi)
         End If
         
         If Len(CStr(rs!depende_2)) = 1 Then
            VAR_DIVISION = "0" + CStr(rs!depende_2)
         Else
            VAR_DIVISION = CStr(rs!depende_2)
         End If
         rsaux2.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + CStr(rs!depende) + "' and vcha_div_division_id = '" + VAR_DIVISION + "' and vcha_sub_subdivision_id = '" + var_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux2.EOF Then
            rsaux.Open "insert into tb_subdivisiones (vcha_tpr_tipo_producto_id, vcha_div_division_id, vcha_sub_subdivision_id, vcha_sub_nombre) values ('" + CStr(rs!depende) + "', '" + VAR_DIVISION + "', '" + var_subdivision + "', '" + CStr(rs!descripcio) + " ')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux2.Close
         rs.MoveNext
   Wend
   rs.Close

   rs.Open "select round(cve_esta,0) as cve_Esta, descripcio from estampad", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If Len(CStr(rs!CVE_ESTA)) = 1 Then
            VAR_ESTAMPADO = "0000" + CStr(rs!CVE_ESTA)
         Else
            If Len(CStr(rs!CVE_ESTA)) = 2 Then
               VAR_ESTAMPADO = "000" + CStr(rs!CVE_ESTA)
            Else
               If Len(CStr(rs!CVE_ESTA)) = 3 Then
                  VAR_ESTAMPADO = "00" + CStr(rs!CVE_ESTA)
               Else
                  If Len(CStr(rs!CVE_ESTA)) = 4 Then
                     VAR_ESTAMPADO = "0" + CStr(rs!CVE_ESTA)
                  Else
                     VAR_ESTAMPADO = CStr(rs!CVE_ESTA)
                  End If
               End If
            End If
         End If
         rsaux2.Open "select * from tb_estampados where vcha_Est_estampado_id = '" + VAR_ESTAMPADO + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux2.EOF Then
            rsaux.Open "insert into tb_estampados (vcha_est_estampado_id, vcha_est_nombre) values ('" + VAR_ESTAMPADO + "', '" + rs!descripcio + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux2.Close
         rs.MoveNext
   Wend
   rs.Close
   MsgBox "termino la migracion de los conceptos", vbOKOnly, "ATENCION"
End Sub

Private Sub COM_LISTA_PRECIOS_Click()
   rs.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) SELECT '01', VCHA_ART_ARTICULO_ID, MONE_ART_PRECIO_BASE FROM TB_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Command1_Click()
    Dim verificador As Integer
    rs.Open "select * from articulos where tipo =  7", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
         var_codigo = Trim(rs!codigo) + "0"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
          rsaux.Open "insert into tb_equivalencias (vcha_art_articulo_id, vcha_equ_codigo_equivalente) values ('" + var_codigo + "','" + Trim(rs!codigo) + "0" + "')", cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cmd_reclasificacion_Click()
   Dim verificador As Integer
   rs.Open "select distinct codigorecl as codigo from articulos", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_codigo = Trim(rs!codigo) + "0"
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
          verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
          If verificador = 10 Then
             verificador = 0
          End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         rsaux2.Open "select * from tb_equivalencias where  vcha_Art_articulo_id = '" + var_codigo + "' and vcha_equ_codigo_equivalente = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux2.EOF Then
            rsaux.Open "insert into tb_equivalencias (vcha_Art_articulo_id, vcha_equ_codigo_equivalente) values ('" + var_codigo + "', '" + rs!codigo + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux2.Close
         rs.MoveNext
   Wend
   rs.Close
   
   
   
   
   
End Sub

Private Sub Command2_Click()
Dim var_año As Integer
Dim var_almacen_Destino As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_numero_causa As Integer
Dim var_elimina As Boolean
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_renglon As Double
Dim verificador As Integer
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   var_empresa = "18"
   var_unidad_organizacional = "16"
   var_almacen_Destino = "AV00125"
   var_clave_movimiento = "EA"
   var_numero_folio = 0
   var_clave_moneda = "1"
   rsaux5.Open "select b.codigo, cantidad_e, a.costo, preciou from mercvistas b, articulos a where allt(cve_agente) = '18' and cantidad_e > 0 and subs(b.codigo,1,10) = a.codigo", var_tabla, adOpenDynamic, adLockOptimistic
   
   var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", "fserna", fun_NombrePc, 0, "", "INVENTARIO INICIAL INICIO S.I.D", "", "", "", "", 0, 0, 0, var_clave_moneda, 0)
   var_numero_folio = var_numero_folio_regreso
   var_primera_vez = False
   While Not rsaux5.EOF
         var_codigo = rsaux5!codigo
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
         verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
         If verificador = 10 Then
            verificador = 0
         End If
          
         var_codigo = var_codigo + Trim(CStr(verificador))
         txt_codigo = var_codigo
         var_costo = rsaux5!costo
         var_precio = rsaux5!preciou
         var_cantidad_leida = rsaux5!cantidad_e
         var_año = 2005
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_año)
            rs.Close
            valor = Trim(txt_codigo)
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
         End If
         rsaux5.MoveNext
   Wend
   Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_inserta = False
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!vcha_emp_empresa_id + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_art_articulo_id + "', " + CStr(rs!floa_ent_cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   var_estatus_movimiento = "I"
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
   var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
   rsaux5.Close
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Dim var_ruta As String
   var_ruta = App.Path
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub
