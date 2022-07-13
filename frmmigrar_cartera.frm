VERSION 5.00
Begin VB.Form frmmigrar_cartera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Migrar cartera"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_completar_folios 
      Caption         =   "Numerar folios de movimientos del almacen"
      Height          =   600
      Left            =   7440
      TabIndex        =   27
      Top             =   6540
      Width           =   2880
   End
   Begin VB.CommandButton Command9 
      Caption         =   "COMPLETAR LA INFORMACION DE LA CARTERA"
      Height          =   585
      Left            =   7470
      TabIndex        =   25
      Top             =   5385
      Width           =   3240
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   720
      Left            =   5325
      TabIndex        =   24
      Top             =   1020
      Width           =   3495
   End
   Begin VB.TextBox txt_limite_credito 
      Height          =   300
      Left            =   3210
      TabIndex        =   15
      Text            =   "0"
      Top             =   7260
      Width           =   1710
   End
   Begin VB.TextBox txt_prioridad 
      Height          =   345
      Left            =   4830
      TabIndex        =   14
      Top             =   6900
      Width           =   2220
   End
   Begin VB.TextBox txt_descuento_3 
      Height          =   330
      Left            =   4830
      TabIndex        =   13
      Text            =   "0"
      Top             =   6540
      Width           =   2205
   End
   Begin VB.TextBox txt_descuento_2 
      Height          =   285
      Left            =   3195
      TabIndex        =   12
      Text            =   "0"
      Top             =   6915
      Width           =   1515
   End
   Begin VB.TextBox txt_descuento_1 
      Height          =   285
      Left            =   3180
      TabIndex        =   11
      Text            =   "0"
      Top             =   6555
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Enabled         =   0   'False
      Height          =   615
      Left            =   345
      TabIndex        =   10
      Top             =   6660
      Width           =   2475
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Abonos "
      Height          =   3555
      Left            =   180
      TabIndex        =   5
      Top             =   2760
      Width           =   10980
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   870
         Left            =   3510
         TabIndex        =   26
         Top             =   2580
         Width           =   2430
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Bonificaciones"
         Height          =   810
         Left            =   195
         TabIndex        =   21
         Top             =   2520
         Width           =   2595
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Bonificaciones financieras"
         Height          =   855
         Left            =   7605
         TabIndex        =   20
         Top             =   1470
         Width           =   2400
      End
      Begin VB.CommandButton Command6 
         Caption         =   "asigna folio de relacion a los descuentos y bonificaciones financieras"
         Height          =   885
         Left            =   2925
         TabIndex        =   19
         Top             =   1500
         Width           =   2370
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Descuentos financieras"
         Height          =   885
         Left            =   5400
         TabIndex        =   17
         Top             =   1470
         Width           =   2100
      End
      Begin VB.CommandButton asignar_establecimientos 
         Caption         =   "Asignar establecimiento"
         Height          =   990
         Left            =   4305
         TabIndex        =   9
         Top             =   210
         Width           =   2220
      End
      Begin VB.CommandButton Command4 
         Caption         =   "aplicar_pagos"
         Height          =   945
         Left            =   8865
         TabIndex        =   8
         Top             =   165
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "se debe correr el procedimiento de genera_descuento"
         Height          =   975
         Left            =   240
         TabIndex        =   18
         Top             =   1470
         Width           =   2625
      End
      Begin VB.Label Label6 
         Caption         =   "AQUI SE DEBE DE CARGA EL PROCEDIMIENTO ALMACENADO DE ""CARGAR_COBRANZA"""
         Height          =   990
         Left            =   6765
         TabIndex        =   16
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   3375
         TabIndex        =   7
         Top             =   495
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   495
         Left            =   2625
         TabIndex        =   6
         Top             =   495
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Cargos "
      Height          =   2445
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   11100
      Begin VB.CommandButton Command1 
         Caption         =   "migrar tablas"
         Height          =   750
         Left            =   600
         TabIndex        =   22
         Top             =   765
         Width           =   2160
      End
      Begin VB.CommandButton cmd_cargos 
         Caption         =   "cargos"
         Height          =   795
         Left            =   2955
         TabIndex        =   1
         Top             =   765
         Width           =   1965
      End
      Begin VB.Label Label8 
         Caption         =   $"frmmigrar_cartera.frx":0000
         Height          =   450
         Left            =   405
         TabIndex        =   23
         Top             =   225
         Width           =   10455
      End
      Begin VB.Label Label1 
         Height          =   405
         Left            =   2445
         TabIndex        =   4
         Top             =   495
         Width           =   1605
      End
      Begin VB.Label Label2 
         Height          =   480
         Left            =   4245
         TabIndex        =   3
         Top             =   405
         Width           =   3240
      End
      Begin VB.Label Label3 
         Caption         =   $"frmmigrar_cartera.frx":009C
         Height          =   660
         Left            =   195
         TabIndex        =   2
         Top             =   1665
         Width           =   10785
      End
   End
End
Attribute VB_Name = "frmmigrar_cartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
   Dim var_serie As String
   Dim var_descuento_aplicar As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_porcentaje_iva As Double
   Dim var_porcentaje_impuesto_2 As Double
   Dim var_porcentaje_impuesto_3 As Double
   Dim var_subimporte As Double
   Dim var_importe_iva As Double
   Dim var_importe_impuesto_2 As Double
   Dim var_importe_impuesto_3 As Double
   Dim var_importe_sin_impuesto As Double
   Dim var_importe_descuento_aplicar As Double
   Dim var_importe_descuento_1 As Double
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_posible_tipo_cambio As Boolean
   Dim var_moneda_local As Integer
   Dim var_tipo_Cambio As Double
   Dim var_importe_total As Double
   Dim var_importe_total_cobranza As Double
   Dim var_importe As Double
   Dim var_numero_nota As Double
   Dim var_cheque As String
   Dim var_importe_saldo As Double
   Dim var_importe_cobranza As Double
   Dim var_descuento_saldo As Double
   Dim var_tipo_documento As String
   Dim var_banco As String
   Dim var_fecha_factura As Date
   Dim i, j As Integer
   Dim var_numero_folio As Double

Private Sub asignar_establecimientos_Click()
     
     rs.Open "select distinct vcha_car_documento, vcha_ser_serie_id, inte_car_numero, vcha_cli_clave_id from tb_encabezado_cartera where vcha_esb_establecimiento_id = 'ESTABLECIMIENTO'", cnn, adOpenDynamic, adLockOptimistic
     If Not rs.EOF Then
        While Not rs.EOF
              rsaux.Open "select vcha_esb_establecimiento_id  from TB_DETALLE_ESTABLECIMIENTOS where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
              If Not rsaux.EOF Then
                rsaux2.Open "update tb_encabezado_cartera set vcha_esb_establecimiento_id = '" + rsaux!vcha_ESB_ESTABLECIMIENTO_id + "' where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "' and inte_Car_numero = " + CStr(rs!inte_Car_numero) + " and vcha_car_documento = '" + rs!vcha_Car_documento + "' and vcha_ser_serie_id = '" + rs!VCHA_SER_SERIE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
              End If
              rsaux.Close
              rs.MoveNext
         Wend
     End If
     rs.Close
End Sub

Private Sub cmd_cargos_Click()
   Dim var_numero As String
   Dim var_serie As String
   Dim var_numero_serie As String
   Dim var_numero_documento As Double
   Dim var_i As Integer, var_j As Integer
   Dim var_documento As String
   Dim var_tipo_documento As String
   Dim var_clase_documento As String
   Dim var_z As Integer
   Dim var_veces As Double
   Dim var_contador As Double
   Dim var_fecha As Date
   Dim var_tipo_Cambio As Double
      Dim var_iva As Double, var_porcentaje_iva As Double, var_porcentaje_descuento_1 As Double, var_porcentaje_descuento_2 As Double, var_descuento_1 As Double, var_descuento_2 As Double
      rs.Open "delete from TB_TEM_MIGRACION_CARTERA", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select count(*) as veces from movsclie", var_tabla, adOpenDynamic, adLockOptimistic
      var_veces = rs!veces
      Me.Label5 = CStr(var_veces)
      Label2.Caption = CStr(var_veces)
      rs.Close
      var_contador = 0
      rs.Open "SELECT * FROM movsclie where allt(tipo) = 'C'", var_tabla, adOpenDynamic, adLockOptimistic
      var_empresa = rs!cveempresa
      While Not rs.EOF
            var_numero = IIf(IsNull(rs!numdocumen), "0", rs!numdocumen)
            If Trim(var_numero) = "1494" Then
               x = 1
            End If
            If IsNumeric(var_numero) Then
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
               var_numero_documento = CDbl(var_numero)
            Else
               var_j = Len(var_numero)
               var_serie = ""
               var_numero_serie = ""
               For var_i = 1 To var_j
                   If Not IsNumeric(Mid(var_numero, var_i, 1)) Then
                      var_serie = var_serie + Mid(var_numero, var_i, 1)
                   Else
                      var_numero_serie = var_numero_serie + Mid(var_numero, var_i, 1)
                   End If
               Next var_i
               If Trim(var_numero_serie) = "" Then
                  var_numero_documento = 1
               Else
                  var_numero_documento = CDbl(var_numero_serie)
               End If
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
            End If
            var_fecha = (rs!fechacaptu)
            var_dia = CStr(Day(var_fecha))
            var_mes = CStr(Month(var_fecha))
            var_año = CStr(Year(var_fecha))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_string = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_tipo_Cambio = 1
            
            var_tipo_Cambio = IIf(IsNull(rs!Dollar), 1, rs!Dollar)
            Cadena = "INSERT INTO TB_TEM_MIGRACION_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_CLI_CLAVE_ID, DTIM_TEM_FECHA_CAPTURA, DTIM_TEM_FECHA_DOCUMENTO, CHAR_TEM_TIPO, VCHA_TEM_CLAVE_DOCUMENTO, VCHA_TEM_SERIE_DOCUMENTO, INTE_TEM_MUMERO_DOCUMENTO, VCHA_AGE_AGENTE_ANTERIOR_ID, VCHA_AGE_AGENTE_ID, VCHA_ZON_ZONA_ANTERIO_ID, VCHA_ZON_ZONA_ID, FLOA_TEM_COMISION, FLOA_TEM_IMPORTE_NETO, DTIM_TEM_FECHA_VENICIMIENTO, VCHA_TEM_REFERENCIA, FLOA_TEM_SALDO_DOCUMENTO, FLOA_TEM_TIPO_CAMBIO)"
            Cadena = Cadena + " values ('" + rs!cveempresa + "', '" + rs!cvecliente + "', '', " + var_fecha_string + ", " + var_fecha_string + ", '" + rs!tipo + "', '" + rs!cvedocumen + "', '" + var_serie + "', " + CStr(var_numero_documento) + ", '" + rs!cveagente + "', '', '" + rs!cvezona + "', '', " + CStr(rs!comision) + ", " + CStr(rs!importenet) + ", " + var_fecha_string + ", '" + rs!Referencia + "', " + CStr(rs!saldodocum) + ", " + CStr(var_tipo_Cambio) + ") "
            rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            
            rsaux.Open "select vcha_cli_clave_id from tb_clientes where vcha_cli_clave_anterior_id = '" + Trim(rs!cvecliente) + "' and vcha_emp_empresa_id = '" + rs!cveempresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_clave_cliente = ""
            If Not rsaux.EOF Then
               var_clave_cliente = rsaux!vcha_cli_clave_id
            End If
            rsaux.Close
            
            rsaux.Open "select distinct vcha_age_agente_id, floa_tpe_iva, vcha_mon_moneda_id, inte_mon_moneda_local from vw_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_clave_agente = ""
            If Not rsaux.EOF Then
               var_clave_agente = IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID)
            End If
            var_tipo_Cambio = 0
            If var_moneda_local = 1 Then
               var_tipo_Cambio = 1
            Else
               var_tipo_Cambio = rs!Dollar
            End If
            var_porcentaje_iva = IIf(IsNull(rsaux!FLOA_TPE_IVA), 0, rsaux!FLOA_TPE_IVA)
            rsaux.Close
            rsaux2.Open "update TB_TEM_MIGRACION_CARTERA set vcha_cli_clave_id = '" + var_clave_cliente + "', vcha_age_agente_id = '" + var_clave_agente + "', FLOA_TEM_PORCENTAJE_IVA = " + CStr(var_porcentaje_iva) + " where vcha_cli_clave_anterior_id = '" + rs!cvecliente + "'", cnn, adOpenDynamic, adLockOptimistic
            
            rs.MoveNext
            var_contador = var_contador + 1
            Me.Refresh
            Me.Label4.Caption = CStr(var_contador)
            Me.Refresh
      Wend
      rs.Close

   
   rs.Open "SELECT cveempresa,round(descuento,2) as descuento, round(dctoptopag,2) as dctoptopag, numfactura , round(iva,2) as iva, round(comision,4) as comision FROM facturas", var_tabla, adOpenKeyset, adLockOptimistic
   var_empresa = rs!cveempresa
   While Not rs.EOF
         var_descuento_1 = 0
         var_descuento_2 = 0
         var_descuento_3 = 0
         var_porcentaje_iva = 0
         var_tipo_Cambio = 1
         var_descuento_1 = IIf(IsNull(rs!descuento), 0, rs!descuento)
         var_descuento_2 = IIf(IsNull(rs!dctoptopag), 0, rs!dctoptopag)
         var_descuento_3 = 0
         'var_tipo_Cambio = IIf(IsNull(rs!comision), 1, rs!comision)
         
         var_numero = IIf(IsNull(rs!NUMFACTURA), "0", rs!NUMFACTURA)
         If IsNumeric(var_numero) Then
            If var_empresa = "02" Then
               var_serie = "X"
            End If
            If var_empresa = "06" Then
               var_serie = "QZ"
            End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
            If var_empresa = "07" Then
               var_serie = "AR"
            End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
            
            var_numero_documento = CDbl(var_numero)
         Else
            var_j = Len(var_numero)
            var_serie = ""
            var_numero_serie = ""
            For var_i = 1 To var_j
                If Not IsNumeric(Mid(var_numero, var_i, 1)) Then
                   var_serie = var_serie + Mid(var_numero, var_i, 1)
                Else
                   var_numero_serie = var_numero_serie + Mid(var_numero, var_i, 1)
                End If
            Next var_i
            var_numero_documento = CDbl(var_numero_serie)
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
         End If
         
         'Cadena = "update TB_TEM_MIGRACION_CARTERA set FLOA_TEM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_TEM_dESCUENTO_2 =  " + CStr(var_descuento_2) + ", FLOA_TEM_TIPO_CAMBIO = " + CStr(var_tipo_Cambio) + " WHERE VCHA_TEM_CLAVE_DOCUMENTO = 'FA' AND VCHA_TEM_SERIE_DOCUMENTO = '" + var_serie + "' AND INTE_TEM_MUMERO_DOCUMENTO = " + CStr(var_numero_documento)
         Cadena = "update TB_TEM_MIGRACION_CARTERA set FLOA_TEM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_TEM_dESCUENTO_2 =  " + CStr(var_descuento_2) + " WHERE VCHA_TEM_CLAVE_DOCUMENTO = 'FA' AND VCHA_TEM_SERIE_DOCUMENTO = '" + var_serie + "' AND INTE_TEM_MUMERO_DOCUMENTO = " + CStr(var_numero_documento)
         rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
'de aqui
   rs.Close
   'se debe de adaptar la tabla de facturas a los campos que tienen decimales a 2 decimales y a un ancho de 7
   
   rs.Open "SELECT cveempresa, numfactura, cvecliente, round(descuento,2) as descuento, round(descuento2,2) as descuento2, round(dctoptopag,2) as dctoptopag, round(comision,4) as comision, round(iva,2) as iva FROM facturas", var_tabla, adOpenDynamic, adLockOptimistic
   var_empresa = rs!cveempresa
   While Not rs.EOF
         var_descuento_1 = 0
         var_descuento_2 = 0
         var_descuento_3 = 0
         var_porcentaje_iva = 0
         var_tipo_Cambio = 1
         var_descuento_1 = IIf(IsNull(rs!descuento), 0, rs!descuento)
         var_descuento_2 = IIf(IsNull(rs!dctoptopag), 0, rs!dctoptopag)
         var_descuento_3 = 0
         var_porcentaje_iva = rs!iva
         var_tipo_Cambio = IIf(IsNull(rs!comision), 1, rs!comision)
        
         var_numero = IIf(IsNull(rs!NUMFACTURA), "0", rs!NUMFACTURA)
         If IsNumeric(var_numero) Then
            If var_empresa = "02" Then
               var_serie = "X"
            End If
            If var_empresa = "06" Then
               var_serie = "QZ"
            End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
            If var_empresa = "07" Then
               var_serie = "AR"
            End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
            
            var_numero_documento = CDbl(var_numero)
         Else
            var_j = Len(var_numero)
            var_serie = ""
            var_numero_serie = ""
            For var_i = 1 To var_j
                If Not IsNumeric(Mid(var_numero, var_i, 1)) Then
                   var_serie = var_serie + Mid(var_numero, var_i, 1)
                Else
                   var_numero_serie = var_numero_serie + Mid(var_numero, var_i, 1)
                End If
            Next var_i
            var_numero_documento = CDbl(var_numero_serie)
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
         End If
         'Cadena = "update TB_TEM_MIGRACION_CARTERA set FLOA_TEM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_TEM_dESCUENTO_2 =  " + CStr(var_descuento_2) + ", FLOA_TEM_PORCENTAJE_IVA = " + CStr(var_porcentaje_iva) + ", FLOA_TEM_TIPO_CAMBIO = " + CStr(var_tipo_Cambio) + " WHERE VCHA_TEM_CLAVE_DOCUMENTO = 'FA' AND VCHA_TEM_SERIE_DOCUMENTO = '" + var_serie + "' AND INTE_TEM_MUMERO_DOCUMENTO = " + CStr(var_numero_documento)
         Cadena = "update TB_TEM_MIGRACION_CARTERA set FLOA_TEM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_TEM_dESCUENTO_2 =  " + CStr(var_descuento_2) + ", FLOA_TEM_PORCENTAJE_IVA = " + CStr(var_porcentaje_iva) + " WHERE VCHA_TEM_CLAVE_DOCUMENTO = 'FA' AND VCHA_TEM_SERIE_DOCUMENTO = '" + var_serie + "' AND INTE_TEM_MUMERO_DOCUMENTO = " + CStr(var_numero_documento)
         rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "Update TB_TEM_MIGRACION_CARTERA set floa_Tem_tipo_cambio = 1 where floa_tem_tipo_cambio = 0", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "Update TB_TEM_MIGRACION_CARTERA set floa_Tem_tipo_cambio = 1 where floa_tem_tipo_cambio is null", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "update TB_TEM_MIGRACION_CARTERA set floa_tem_porcentaje_iva =  0 where floa_tem_porcentaje_iva is null", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "update tb_tem_migracion_cartera set floa_tem_descuento_1 = 0 where floa_tem_descuento_1 is null", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "update tb_tem_migracion_Cartera set floa_tem_descuento_2 = 0 where floa_tem_Descuento_2 is null", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select * from tb_tem_migracion_cartera ", cnn, adOpenDynamic, adLockOptimistic
   Label4 = CStr(rs.Fields.Count)
   var_i = 0
   While Not rs.EOF
         var_documento = rs!VCHA_TEM_CLAVE_DOCUMENTO
         If var_documento = "NC" Then
            var_tipo_documento = "NG"
            var_clase_documento = "SM"
         Else
            If var_documento = "CH" Then
               var_tipo_documento = "CH"
               var_clase_documento = "CH"
            Else
               var_tipo_documento = "FA"
               var_clase_documento = "FA"
            End If
         End If
         If rsaux.State Then
            rsaux.Close
         End If
         rsaux4.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_ser_serie_id = '" + rs!VCHA_TEM_SERIE_DOCUMENTO + "' and  vcha_car_tipo_documento = '" + var_tipo_documento + "' and  inte_car_numero = " + CStr(rs!INTE_TEM_MUMERO_DOCUMENTO), cnn, adOpenDynamic, adLockOptimistic
         If rsaux4.EOF Then
            rsaux.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_importe_neto = 0
               var_porcentaje_descuento_1 = 0
               var_porcentaje_descuento_2 = 0
               var_descuento_1 = 0
               var_descuento_2 = 0
               var_porcentaje_iva = 0
               var_iva = 0
               var_porcentaje_descuento_1 = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_1), 0, rs!FLOA_TEM_DESCUENTO_1)
               var_porcentaje_descuento_2 = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_2), 0, rs!FLOA_TEM_DESCUENTO_2)
               var_porcentaje_iva = IIf(IsNull(rs!FLOA_TEM_PORCENTAJE_IVA), 0, rs!FLOA_TEM_PORCENTAJE_IVA)
               var_iva = rs!FLOA_TEM_IMPORTE_NETO - (rs!FLOA_TEM_IMPORTE_NETO / (1 + (var_porcentaje_iva / 100)))
               var_subimporte = rs!FLOA_TEM_IMPORTE_NETO - var_iva
               var_descuento_2 = (var_subimporte * 100) / (100 - var_porcentaje_descuento_2)
               var_descuento_2 = var_descuento_2 - var_subimporte
               var_descuento_1 = ((var_subimporte + var_descuento_2) * 100) / (100 - var_porcentaje_descuento_1)
               var_descuento_1 = var_descuento_1 - (var_subimporte + var_descuento_2)
               var_importe_neto = var_subimporte + var_descuento_1 + var_descuento_2
               Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2,"
               Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE,FLOA_CAR_IMPORTE_NETO,VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO,VCHA_AUD_MAQUINA,VCHA_AUD_FECHA,FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS) values"
               Cadena = Cadena + "('" + rs!VCHA_EMP_EMPRESA_ID + "', '', '" + var_tipo_documento + "', '" + var_documento + "', '" + var_clase_documento + "', "
               Cadena = Cadena + CStr(rs!INTE_TEM_MUMERO_DOCUMENTO) + ",'+', '', '', 0, '" + CStr(rs!DTIM_TEM_FECHA_DOCUMENTO) + "', '"
               Cadena = Cadena + IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID)
               Cadena = Cadena + "', '" + IIf(IsNull(rsaux!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux!VCHA_GAC_GRUPO_aCTUAL_ID) + "', '" + IIf(IsNull(rsaux!vcha_gre_grupo_real_id), "", rsaux!vcha_gre_grupo_real_id) + "', '" + rsaux!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "','ESTABLECIMIENTO', " + CStr(IIf(IsNull(rsaux!inte_pla_dias), 30, rsaux!inte_pla_dias)) + "," + CStr(rs!FLOA_TEM_PORCENTAJE_IVA) + ", 0, 0, " + CStr(rs!FLOA_TEM_DESCUENTO_1) + ", " + CStr(rs!FLOA_TEM_DESCUENTO_2) + ", 0, " + CStr(var_importe_neto) + ", " + CStr(var_iva) + ", 0, 0, " + CStr(var_descuento_1) + ", " + CStr(var_descuento_2)
               Cadena = Cadena + ",0, " + CStr(var_subimporte) + ", " + CStr(rs!FLOA_TEM_IMPORTE_NETO) + ", '', 'USUARIO', 'MAQUINA', " + CStr(rs!DTIM_TEM_FECHA_CAPTURA) + ", 0,null, null, "
               Cadena = Cadena + IIf(IsNull(rsaux!vcha_mon_moneda_id), "1", rsaux!vcha_mon_moneda_id) + ", " + CStr(rs!FLOA_TEM_TIPO_CAMBIO) + ", '" + rs!VCHA_TEM_SERIE_DOCUMENTO + "', 'I')"
               rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If rsaux3.State = 1 Then
                  rsaux3.Close
               End If
               rsaux3.Open "INSERT INTO TB_ESTADO_CUENTA (VCHA_EMP_EMPRESA_ID, VCHA_ECU_SERIE_CARGO, VCHA_ECU_MOVIMIENTO_CARGO, INTE_ECU_NUMERO_CARGO, FLOA_ECU_IMPORTE_CARGO, FLOA_ECU_IMPORTE_ABONO) Values ('" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!VCHA_TEM_SERIE_DOCUMENTO + "', '" + var_documento + "'," + CStr(rs!INTE_TEM_MUMERO_DOCUMENTO) + " , " + CStr(rs!FLOA_TEM_IMPORTE_NETO) + ", 0)", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
         End If
         rsaux4.Close
         var_i = var_i + 1
         Label2 = CStr(var_i)
         Me.Refresh
         rs.MoveNext
   Wend
   rs.Close
   'hasta aqui
MsgBox "Se a terminado la migración de la cartera", vbOKOnly, "ATENCION"
 
   
End Sub

Private Sub cmd_completar_folios_Click()
   rs.Open "select vcha_emp_empresa_id, vcha_uor_unidad_id, max(inte_sal_mumero) from tb_salidas group by vcha_emp_empresa_id, vcha_uor_unidad_id"
End Sub

Private Sub Command1_Click()
    Dim var_num_bo As Double
    Dim var_num_dv As Double
    Dim var_num_ca As Double
    Dim var_num_ds As Double
    
    rs.Open "delete from tb_temp_movsclie ", cnn, adOpenDynamic, adLockOptimistic
    rs.Open "delete from tb_temp_movsabon ", cnn, adOpenDynamic, adLockOptimistic
    'rs.Open "update tb_tem_movsclie set cveagente = '03' where allt(cveagente) =  '07'", var_tabla, adOpenDynamic, adLockOptimistic
    'rs.Open "update tb_tem_movsclie set cveagente = '09' where allt(cveagente) =  '10'", var_tabla, adOpenDynamic, adLockOptimistic
    'rs.Open "update tb_tem_movsclie set cveagente = '09' where allt(cveagente) =  '11'", var_tabla, adOpenDynamic, adLockOptimistic
    'rs.Open "update tb_tem_movsclie set cveagente = '05' where allt(cveagente) =  '47'", var_tabla, adOpenDynamic, adLockOptimistic
    'rs.Open "update tb_tem_movsclie set cveagente = '38' where allt(cveagente) =  '31'", var_tabla, adOpenDynamic, adLockOptimistic
    rs.Open "SELECT * FROM movsclie where allt(tipo) = 'A'", var_tabla, adOpenDynamic, adLockOptimistic
    var_conta = 0
    rsaux.Open "select max(inte_car_numero) FROM TB_ENCABEZADO_CARTERA where vcha_car_clase_id = 'CA'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       var_num_ca = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
    Else
       var_num_ca = 1
    End If
    rsaux.Close
    rsaux.Open "select max(inte_car_numero) FROM TB_ENCABEZADO_CARTERA  where vcha_car_clase_id = 'BO'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       var_num_bo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
    Else
       var_num_bo = 1
    End If
    rsaux.Close
    
    rsaux.Open "select max(inte_car_numero) FROM TB_ENCABEZADO_CARTERA where vcha_car_clase_id = 'DV'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       var_num_dv = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
    Else
       var_num_dv = 1
    End If
    rsaux.Close
    rsaux.Open "select max(inte_car_numero) FROM TB_ENCABEZADO_CARTERA  where vcha_car_clase_id = 'DS'", cnn, adOpenDynamic, adLockOptimistic
    If Not rsaux.EOF Then
       var_num_ds = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
    Else
       var_num_ds = 1
    End If
    rsaux.Close
    While Not rs.EOF
            var_conta = var_conta + 1
            Label4 = CStr(var_conta)
            Me.Refresh
            var_numero = IIf(IsNull(rs!numdocumen), "0", rs!numdocumen)
            var_empresa = rs!cveempresa
            If IsNumeric(var_numero) Then
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
               var_numero_documento = CDbl(var_numero)
            Else
               var_j = Len(var_numero)
               var_serie = ""
               var_numero_serie = ""
               For var_i = 1 To var_j
                   If Not IsNumeric(Mid(var_numero, var_i, 1)) Then
                      var_serie = var_serie + Mid(var_numero, var_i, 1)
                   Else
                      var_numero_serie = var_numero_serie + Mid(var_numero, var_i, 1)
                   End If
               Next var_i
               var_numero_documento = CDbl(var_numero_serie)
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
            End If
            
            
            var_numero_abono = IIf(IsNull(rs!numerodcto), "0", rs!numerodcto)
            If IsNumeric(var_numero_abono) Then
               If var_empresa = "02" Then
                  var_serie_abono = "X"
               End If
               If var_empresa = "06" Then
                  var_serie_abono = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie_abono = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie_abono = "AR"
               End If
               If var_empresa = "15" Then
                  var_serie_abono = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie_abono = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie_abono = "AA"
               End If
               var_numero_documento_abono = CDbl(var_numero_abono)
            Else
               var_j = Len(var_numero_abono)
               var_serie_abono = ""
               var_numero_serie_abono = ""
               For var_i = 1 To var_j
                   If Not IsNumeric(Mid(var_numero_abono, var_i, 1)) Then
                      var_serie_abono = var_serie_abono + Mid(var_numero_abono, var_i, 1)
                   Else
                      var_numero_serie_abono = var_numero_serie_abono + Mid(var_numero_abono, var_i, 1)
                   End If
               Next var_i
               If Trim(var_numero_serie_abono) <> "" Then
                  var_numero_documento_abono = CDbl(var_numero_serie_abono)
               Else
                  var_numero_documento_abono = 1
               End If
               If var_empresa = "06" Then
                  var_serie_abono = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie_abono = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie_abono = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie_abono = "X"
               End If
               If var_empresa = "15" Then
                  var_serie_abono = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie_abono = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie_abono = "AA"
               End If
            End If
            
            
            var_fecha = (rs!fechacaptu)
            var_dia = CStr(Day(var_fecha))
            var_mes = CStr(Month(var_fecha))
            var_año = CStr(Year(var_fecha))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_string = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
          If rs!cvedocumen = "BO" Then
             nUMERO_DOCUMENTO = var_num_bo
             var_num_bo = var_num_bo + 1
          Else
             If rs!cvedocumen = "DS" Then
                nUMERO_DOCUMENTO = var_num_ds
                var_num_ds = var_num_ds + 1
             Else
                If rs!cvedocumen = "CA" Then
                   nUMERO_DOCUMENTO = var_num_ca
                   var_num_ca = var_num_ca + 1
                Else
                   If rs!cvedocumen = "DV" Then
                      nUMERO_DOCUMENTO = var_num_dv
                      var_num_dv = var_num_dv + 1
                   Else
                      nUMERO_DOCUMENTO = var_numero_documento_abono
                   End If
                End If
             End If
          End If
          
          Cadena = "INSERT INTO TB_TEMP_MOVSCLIE (VCHA_EMP_EMPRESA_ID, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_CLI_CLAVE_ID, DTIM_MVS_FECHA_DOCUMENTO, VCHA_MVS_CLAVE_DOCUMENTO, VCHA_MVS_NUMERO_DOCUMENTO,VCHA_MVS_SERIE, INTE_MVS_NUMERO, FLOA_MVS_IMPORTE_NETO, VCHA_MVS_PARTIDA,VCHA_MVS_CLAVE_DOCUMENTO_ABONO, VCHA_MVS_SERIE_DOCUMENTO_ABONO, INTE_MVS_NUMERO_DOCUMENTO_ABONO, FLOA_MVS_IMPORTE_ABONO, VCHA_AGE_AGENTE_ANTERIOR_ID, floa_mvs_tipo_cambio) values"
          Cadena = Cadena + "( '" + rs!cveempresa + "', '" + rs!cvecliente + "',  '', " + var_fecha_string + ", '" + rs!cvedocumen + "', '" + rs!numdocumen + "', '" + var_serie + "', " + CStr(nUMERO_DOCUMENTO) + ",  " + CStr(rs!importenet) + ",  '" + rs!PARTIDA + "', '" + rs!CLAVEDCTO + "', '" + var_serie_abono + "',  " + CStr(var_numero_documento_abono) + ", 0, '" + rs!cveagente + "', " + CStr(rs!Dollar) + ")"
          rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
    var_conta = 0
    z = 1
    If z = 1 Then
    rs.Open "SELECT * FROM MOVSABON", var_tabla, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          var_conta = var_conta + 1
          Label5 = CStr(var_conta)
          Me.Refresh
          var_numero = IIf(IsNull(rs!numdocumen), "0", rs!numdocumen)
          var_empresa = rs!cveempresa
          If IsNumeric(var_numero) Then
             If var_empresa = "02" Then
                var_serie = "X"
             End If
             If var_empresa = "06" Then
                var_serie = "QZ"
             End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
             If var_empresa = "07" Then
                var_serie = "AR"
             End If
             If var_empresa = "15" Then
                var_serie = "ER"
             End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
             var_numero_documento = CDbl(var_numero)
          Else
             var_j = Len(var_numero)
             var_serie = ""
             var_numero_serie = ""
             For var_i = 1 To var_j
                 If Not IsNumeric(Mid(var_numero, var_i, 1)) Then
                    var_serie = var_serie + Mid(var_numero, var_i, 1)
                 Else
                    var_numero_serie = var_numero_serie + Mid(var_numero, var_i, 1)
                 End If
             Next var_i
             var_numero_documento = CDbl(var_numero_serie)
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
          End If
            var_numero_abono = IIf(IsNull(rs!numerodcto), "0", rs!numerodcto)
            If IsNumeric(var_numero_abono) Then
               If var_empresa = "02" Then
                  var_serie_abono = "X"
               End If
               If var_empresa = "06" Then
                  var_serie_abono = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie_abono = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie_abono = "AR"
               End If
               If var_empresa = "15" Then
                  var_serie_abono = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie_abono = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie_abono = "AA"
               End If
               
               var_numero_documento_abono = CDbl(var_numero_abono)
            Else
               var_j = Len(var_numero_abono)
               var_serie_abono = ""
               var_numero_serie_abono = ""
               For var_i = 1 To var_j
                   If Not IsNumeric(Mid(var_numero_abono, var_i, 1)) Then
                      var_serie_abono = var_serie_abono + Mid(var_numero_abono, var_i, 1)
                   Else
                      var_numero_serie_abono = var_numero_serie_abono + Mid(var_numero_abono, var_i, 1)
                   End If
               Next var_i
               var_numero_documento_abono = CDbl(var_numero_serie_abono)
               If var_empresa = "06" Then
                  var_serie = "QZ"
               End If
               If var_empresa = "17" Then
                  var_serie = "BI"
               End If
               If var_empresa = "07" Then
                  var_serie = "AR"
               End If
               If var_empresa = "02" Then
                  var_serie = "X"
               End If
               If var_empresa = "15" Then
                  var_serie = "ER"
               End If
               If var_empresa = "16" Then
                  var_serie = "MG"
               End If
               If var_empresa = "18" Then
                  'var_serie = "AA"
               End If
            End If
          Cadena = "INSERT INTO TB_TEMP_MOVSABON (VCHA_EMP_EMPRESA_ID, VCHA_MVA_CLAVE_ABONO, VCHA_MVA_NUMERO_ABONO, VCHA_MVA_CLAVE_DOCUMENTO, VCHA_MVA_NUMERO_DOCUMENTO, VCHA_MVA_CLAVE_DOCUMENTO_ABONO,VCHA_MVA_SERIE_ABONO ,INTE_MVA_NUMERO_DOCUMENTO_ABONO, FLOA_MVA_DESCUENTO,FLOA_MVA_DESCUENTO_CAPTURADO, VCHA_MVA_PARTIDA, VCHA_MVA_PARTIDA_ABONO) Values"
          Cadena = Cadena + "('" + rs!cveempresa + "','" + rs!CVEABONO + "','" + rs!NUMABONO + "', '" + rs!cvedocumen + "','" + rs!numdocumen + "','" + rs!CLAVEDCTO + "', '" + var_serie_abono + "'," + CStr(var_numero_documento_abono) + "," + CStr(rs!DCTOAPLIC) + ", " + CStr(rs!DCTOCAPTUR) + ", '" + rs!PARTIDA + "', '" + rs!PARTIABONO + "')"
          rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
    End If
End Sub

Private Sub Command2000_Click()
   rs.Open "select * from VW_TEMP_RELACION_COBRANZA", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_empresa_cobranza = rs!VCHA_EMP_EMPRESA_ID
         var_cliente = rs!vcha_cli_clave_id
         var_agente = rs!VCHA_AGE_AGENTE_ID
         var_cheque = "EFVO"
         var_fecha = rs!DTIM_MVS_FECHA_DOCUMENTO
         var_fecha_cheque = "EFVO"
         var_importe = rs!floa_mvs_importe_neto
         var_descuento = rs!floa_mva_descuento
         var_factura = rs!inte_mva_numero_documento_abono
         var_serie = rs!vcha_mvA_serie_abono
         var_tipo = rs!VCHA_MVA_CLAVE_DOCUMENTO_ABONO
         var_folio = rs!vcha_Rco_folio
         var_banco = "EFVO"
         var_cero = 0
         var_partida = CInt(rs!VCHA_MVA_PARTIDA)
         rsaux.Open "EXECUTE RELACION_COBRANZA_I '" + var_empresa_cobranza + "', '08', '" + var_folio + "', '" + CStr(rs!DTIM_MVS_FECHA_DOCUMENTO) + "', '" + var_agente + "', '" + var_cliente + "', 'EFVO', '" + CStr(rs!DTIM_MVS_FECHA_DOCUMENTO) + "', " + CStr(var_importe) + ", " + CStr(var_descuento) + ", " + CStr(var_factura) + ", 0, 0, " + CStr(var_partida) + ", 0, '" + var_serie + "', '" + var_tipo + "', 'EFVO'", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
End Sub

Private Sub Command10_Click()
   Set TB_GRUPOSACTUALES = New TB_GRUPOSACTUALES
   Set TB_GRUPOSREALES = New TB_GRUPOSREALES
   Set TB_TITULARES = New TB_TITULARES
   Set TB_CLIENTES = New TB_CLIENTES
   Set TB_ESTABLECIMIENTOS = New TB_ESTABLECIMIENTOS
   
   Dim var_clave_titular As String
   Dim var_clave_grupo_actual As String
   Dim var_clave_grupo_real As String
   Dim var_clave_establecimiento As String
   Dim var_agente As String
   Dim var_prioridad As String
   
   Dim var_clave_titular_nueva As String
   Dim var_clave_grupo_actual_nueva As String
   Dim var_clave_grupo_real_nueva As String
   Dim var_clave_establecimiento_nueva As String
   Dim var_clave_cliente_nueva As String
   Dim var_agente_nueva As String
   Dim var_nombre_cliente As String
   Dim var_direccion As String
   Dim var_cp As String
   Dim descuento_1 As Double
   Dim descuento_2 As Double
   Dim limite_credito As Double
   Dim var_plazo As Double
   Dim var_clave_plazo As String
   
   Dim var_lista As String
   Dim var_moneda As String
   Dim var_tipo_cliente As String
   
   var_ruta = App.Path
   Set var_tabla = Nothing
   Set var_tabla = CreateObject("ADODB.connection")
   var_ruta = "c:\sistemas\desarrollo\integral\clientes\"
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   cnn.BeginTrans
   rs.Open "select cvecliente, fechaalta, rfc from clientes", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux4.Open "select * from tb_clientes where vcha_cli_clave_anterior_id = '" + Trim(rs!cvecliente) + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux4.EOF Then
            rsaux3.Open "select cvecliente , razonsocia, descuento,dctoptopag,cvetitular,cvefamdcto,cvefamcte,cveagente,direccion,plazopagar, tipoclient, codigopost from clientes where cvecliente = '" + rs!cvecliente + "'", var_tabla, adOpenDynamic, adLockOptimistic
            var_nombre_cliente = IIf(IsNull(rsaux3!razonsocia), "", rsaux3!razonsocia)
            var_clave_titular = IIf(IsNull(rsaux3!cvetitular), "", rsaux3!cvetitular)
            var_clave_grupo_real = IIf(IsNull(rsaux3!cvefamdcto), "", rsaux3!cvefamdcto)
            var_clave_grupo_actual = IIf(IsNull(rsaux3!cvefamcte), "", rsaux3!cvefamcte)
            var_agente = IIf(IsNull(rsaux3!cveagente), "", rsaux3!cveagente)
            var_direccion = IIf(IsNull(rsaux3!DIRECCION), "", rsaux3!DIRECCION)
            var_cp = IIf(IsNull(rsaux3!CODIGOPOST), "", rsaux3!CODIGOPOST)
            descuento_1 = IIf(IsNull(rsaux3!descuento), 0, rsaux3!descuento)
            descuento_2 = IIf(IsNull(rsaux3!dctoptopag), 0, rsaux3!dctoptopag)
            var_plazo = IIf(IsNull(rsaux3!plazopagar), 0, rsaux3!plazopagar)
            var_prioridad = IIf(IsNull(rsaux3!tipoclient), "z", rsaux3!tipoclient)
            rsaux3.Close
            var_lista = "01"
            var_moneda = "1"
            var_tipo_cliente = "M"
            
            rsaux2.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ANTERIOR_ID = '" + var_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_agente = IIf(IsNull(rsaux2!VCHA_AGE_AGENTE_ID), "00076", rsaux2!VCHA_AGE_AGENTE_ID)
            Else
               var_agente = "00076"
            End If
            rsaux2.Close
            rsaux2.Open "select * from tb_plazos where inte_pla_dias = " + CStr(var_plazo), cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_clave_plazo = IIf(IsNull(rsaux2!VCHA_PLA_PLAZO_ID), "2", rsaux2!VCHA_PLA_PLAZO_ID)
            Else
               var_clave_plazo = "2"
            End If
            rsaux2.Close
            rsaux2.Open "select * from detatien where cvecliente = '" + rs!cvecliente + "'", var_tabla, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_clave_establecimiento = IIf(IsNull(rsaux2!cvetienda), "", rsaux2!cvetienda)
            Else
               var_clave_establecimiento = ""
            End If
            rsaux2.Close
            If var_clave_grupo_actual <> "" Then
               rsaux2.Open "select * from tb_gruposactuales where vcha_gac_grupo_actual_anterior_id = '" + var_clave_grupo_actual + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_clave_grupo_actual_nueva = rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID
               Else
                  ok = TB_GRUPOSACTUALES.Anadir("", var_nombre_cliente, CStr(descuento_1), CStr(descuento_2))
                  var_clave_grupo_actual_nueva = var_grupo_actual_regreso
               End If
               rsaux2.Close
            Else
               ok = TB_GRUPOSACTUALES.Anadir("", var_nombre_cliente, descuento_1, descuento_2)
               var_clave_grupo_actual_nueva = var_grupo_actual_regreso
            End If
            rsaux2.Open "update tb_gruposactuales set vcha_gac_grupo_actual_anterior_id = '" + var_clave_grupo_actual_id + "' where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual_nueva + "'", cnn, adOpenDynamic, adLockOptimistic
            If Trim(var_clave_grupo_real) <> "" Then
               rsaux2.Open "SELECT * FROM TB_GRUPOSREALES WHERE VCHA_GRE_GRUPO_REAL_ANTERIOR_ID = '" + Trim(var_clave_grupo_real) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_clave_grupo_real_nueva = rsaux2!vcha_gre_grupo_real_id
               Else
                  ok = TB_GRUPOSREALES.Anadir(var_clave_grupo_actual_nueva, "x", Trim(var_nombre_cliente), descuento_1, descuento_2, 0, var_prioridad)
                  var_clave_grupo_real_nueva = var_grupo_real_regreso
               End If
               rsaux2.Close
            Else
              ok = TB_GRUPOSREALES.Anadir(var_clave_grupo_actual_nueva, "x", Trim(var_nombre_cliente), descuento_1, descuento_2, 0, var_prioridad)
              var_clave_grupo_real_nueva = var_grupo_real_regreso
            End If
            rsaux2.Open "update tb_gruposreales set vcha_gre_grupo_real_anterior_id = '" + var_clave_grupo_real + "' where vcha_gre_grupo_real_id = '" + var_clave_grupo_real_nueva + "'", cnn, adOpenDynamic, adLockOptimistic
            If Trim(var_clave_titular) <> "" Then
               rsaux2.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + Trim(var_clave_titular) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_clave_titular_nueva = rsaux2!vcha_tit_titular_id
               Else
                  ok = TB_TITULARES.Anadir(var_clave_grupo_real_nueva, "", var_nombre_cliente, "", "", "", "", "", var_direccion, "", limite_credito, var_cp)
                  var_clave_titular_nueva = var_titular_regreso
               End If
               rsaux2.Close
            Else
               ok = TB_TITULARES.Anadir(var_clave_grupo_real_nueva, "", var_nombre_cliente, "", "", "", "", "", var_direccion, "", limite_credito, var_cp)
               var_clave_titular_nueva = var_titular_regreso
            End If
            rsaux2.Open "update tb_titulares set vcha_tit_titular_anterior_id = '" + var_clave_titular + "' where vcha_tit_titular_id = '" + var_clave_titular_nueva + "'", cnn, adOpenDynamic, adLockOptimistic
            ok = TB_CLIENTES.Anadir("", var_nombre_cliente, var_nombre_cliente, CStr(IIf(IsNull(rs!fechaalta), Date, rs!fechaalta)), var_agente, "", "", IIf(IsNull(rs!rfc), "", rs!rfc), var_moneda, var_clave_plazo, var_tipo_cliente, var_lista, "", "", "", 1, var_clave_titular_nueva, var_prioridad, "", "", "", "", "", var_direccion, var_cp, "", 0, 1, rs!cvecliente)
            var_clave_cliente_nueva = var_cliente_regreso
            rsaux2.Open "select * from tb_establecimientoS where vcha_esb_establecimiento_anterior_id = '" + var_clave_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_clave_establecimiento_nueva = IIf(IsNull(rsaux2!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux2!vcha_ESB_ESTABLECIMIENTO_id)
            Else
              ok = TB_ESTABLECIMIENTOS.Anadir(var_clave_titular_nueva, "", var_nombre_cliente, "", "", "", "", var_direccion, "", "", "", var_cp)
              var_clave_establecimiento_nueva = var_establecimiento_regreso
              rsaux.Open "insert into TB_DETALLE_ESTABLECIMIENTOS (vcha_esb_establecimiento_id, vcha_cli_clave_id) values( '" + var_clave_establecimiento_nueva + "', '" + var_clave_cliente_nueva + "')", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux2.Close
            rsaux2.Open "update tb_establecimientos set vcha_esb_establecimiento_anterior_id = '" + var_clave_establecimiento + "' where vcha_esb_establecimiento_id = '" + var_clave_establecimiento_nueva + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux4.Close
         rs.MoveNext
   Wend
   rs.Close
   cnn.CommitTrans
   cnn.BeginTrans
   rs.Open "select * from tb_clientes", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_clave_cliente = rs!vcha_cli_clave_id
         var_clave_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         var_direccion = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
         var_cp = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
         limite_credito = 0
         rsaux2.Open "select * from tb_titulares where vcha_tit_titular_id = '" + var_clave_titular + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux2.EOF Then
            ok = TB_GRUPOSACTUALES.Anadir("", var_nombre_cliente, 0, 0)
            ok = TB_GRUPOSREALES.Anadir(var_grupo_actual_regreso, "x", Trim(var_nombre_cliente), 0, 0, 0, "z")
            ok = TB_TITULARES.Anadir(var_grupo_real_regreso, "", var_nombre_cliente, "", "", "", "", "", var_direccion, "", limite_credito, var_cp)
            rsaux.Open "update tb_clientes set vcha_tit_titular_id = '" + var_titular_regreso + "' where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         Else
            var_clave_grupo_real = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
            var_nombre_cliente = IIf(IsNull(rsaux2!VCHA_TIT_NOMBRE), "", rsaux2!VCHA_TIT_NOMBRE)
            rsaux.Open "select * from TB_GRUPOSREALES where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               ok = TB_GRUPOSACTUALES.Anadir("", var_nombre_cliente, 0, 0)
               ok = TB_GRUPOSREALES.Anadir(var_grupo_actual_regreso, "x", Trim(var_nombre_cliente), 0, 0, 0, "z")
            Else
               var_clave_grupo_actual = IIf(IsNull(rsaux!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux!VCHA_GAC_GRUPO_aCTUAL_ID)
               var_nombre_GRUPO_ACTUAL = IIf(IsNull(rsaux!vcha_gre_nombre), "", rsaux!vcha_gre_nombre)
               rsaux3.Open "select * from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux3.EOF Then
                  ok = TB_GRUPOSACTUALES.Anadir("", CStr(var_nombre_GRUPO_ACTUAL), CStr(0), CStr(0))
               End If
               rsaux3.Close
            End If
            rsaux.Close
         End If
         rsaux2.Close
         var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
         If Trim(var_agente) = "" Then
            rsaux.Open "update tb_clientes set vcha_age_agente_id = '00076' where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         Else
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + var_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               rsaux2.Open "update tb_clientes set vcha_age_agente_id = '00076' where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux.Close
         End If
         var_tipo_cliente = rs!VCHA_TCL_TIPO_CLIENTE_ID
         If var_tipo_cliente = "E" Then
            var_lista = "02"
            var_clave_lista = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
            If var_clave_lista = "" Then
               rsaux2.Open "update tb_clientes set vcha_lis_lista_id = '" + var_lista + "' where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
         If var_tipo_cliente = "M" Then
            var_lista = "01"
            var_clave_lista = IIf(IsNull(rs!vcha_LIS_LISTA_iD), "", rs!vcha_LIS_LISTA_iD)
            If var_clave_lista = "" Then
               rsaux2.Open "update tb_clientes set vcha_lis_lista_id = '" + var_lista + "' where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
         rs.MoveNext
   Wend
   rs.Close
   cnn.CommitTrans
End Sub

Private Sub Command11_Click()
    rs.Open "select * from tb_clientes", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + IIf(IsNull(rs!VCHA_RUT_RUTA_ID), "", rs!VCHA_RUT_RUTA_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
          If rsaux.EOF Then
             rsaux2.Open "select * from tb_rutas where vcha_age_agente_id = '" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux2.EOF Then
                rsaux3.Open "UPDATE TB_CLIENTES SET VCHA_RUT_RUTA_ID = '" + rsaux2!VCHA_RUT_RUTA_ID + "' WHERE VCHA_CLI_CLAVE_ID = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
             Else
                rsaux3.Open "UPDATE TB_CLIENTES SET VCHA_RUT_RUTA_ID = '0069' WHERE VCHA_CLI_CLAVE_ID = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
             End If
             rsaux2.Close
          End If
          rsaux.Close
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Command12_Click()
    cnn.BeginTrans
    rs.Open "delete from tb_encabezado_cartera where vcha_car_clase_id = 'BF'", cnn, adOpenDynamic, adLockOptimistic
    rs.Open "select * from tb_estado_cuenta where vcha_ecu_movimiento_abono = 'DF'"
    While Not rs.EOF
          rsaux.Open "update tb_saldos set floa_sal_importe = floa_sal_importe + " + CStr(rs!floa_ecu_importe_abono) + " where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_ser_serie_id = '" + rs!vcha_ecu_serie_cargo + "' and inte_car_numero = " + CStr(rs!inte_ecu_numero_cargo), cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
    rs.Open "delete from tb_estado_cuenta where vcha_ecu_movimiento_abono = 'DF'", cnn, adOpenDynamic, adLockOptimistic
    cnn.CommitTrans
End Sub
Private Sub Command2_Click()
   Dim var_tolerancia_saldo As Integer
   Dim var_descuento_otorgado As Double
   Dim var_descuento_aplicado As Double
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_contador_lineas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim si, i, n As Integer
   Dim var_saldo As Double
   Dim var_serie_cargo As String
   Dim var_numero_nota_inicio As Integer
   Dim var_factura As Double
   Dim var_k As Integer
   Dim var_descuentos As Integer
   Dim var_desc_otorgado_str As String
   Dim var_desc_apilcado_str As String
   Dim var_documento As String
   Dim var_fecha_documento As Date
   Set TB_ENCABEZADO_CARTERA_MIGRACION = New TB_ENCABEZADO_CARTERA_MIGRACION
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   cnn.BeginTrans
   var_tolerancia_saldo = 6
   x = 1
   If x = 1 Then
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   rsaux4.Open "SELECT * FROM VW_TEMP_RELACION_COBRANZA_SUMATORIA_DF", cnn, adOpenDynamic, adLockOptimistic
   var_i = 0
   While Not rsaux4.EOF
         If rs.State = 1 Then
            rs.Close
         End If
         'rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + rsaux4!vcha_emp_empresa_id + "' and vcha_car_tipo_documento = 'NC' and vcha_car_documento = 'DF' and vcha_car_clase_id = 'DF' and inte_car_numero = " + CStr(rsaux4!inte_mvs_numero) + " and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         'If rs.EOF Then
            var_i = var_i + 1
            Label4 = CStr(var_i)
            Me.Refresh
            Label4.Refresh
            var_almacen = ""
            var_fecha_documento = rsaux4!DTIM_MVS_FECHA_DOCUMENTO
            var_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
            var_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
            var_titular = rsaux4!vcha_tit_titular_id
            var_agente = IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID)
            var_cliente = rsaux4!vcha_cli_clave_id
            var_establecimiento = IIf(IsNull(rsaux4!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux4!vcha_ESB_ESTABLECIMIENTO_id)
            If rsaux2.State = 1 Then
                rsaux2.Close
            End If
            var_iva = IIf(IsNull(rsaux4!FLOA_TPE_IVA), 0, rsaux4!FLOA_TPE_IVA)
            var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
            var_tipo_Cambio = IIf(IsNull(rsaux4!floa_MVS_TIPO_cambio), 1, rsaux4!floa_MVS_TIPO_cambio)
            If var_tipo_Cambio = 0 Then
               var_tipo_Cambio = 1
            End If
         
            var_serie = rsaux4!vcha_mvs_serie
            var_numero_nota = CDbl(rsaux4!vcha_mvs_numero_documento)
            var_importe = rsaux4!importe_neto
            var_subimporte = var_importe / (1 + (var_iva / 100))
            var_importe_iva = var_importe - var_subimporte
            var_insertar = TB_ENCABEZADO_CARTERA_MIGRACION.Anadir(rsaux4!VCHA_EMP_EMPRESA_ID, "08", "NC", "DF", "DF", var_numero_nota, "-", "", "", 0, CStr(var_fecha_documento), var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe * var_tipo_Cambio, var_importe_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_subimporte * var_tipo_Cambio, var_importe * var_tipo_Cambio, "", var_clave_usuario_global, "", var_fecha_documento, 0, var_fecha_documento, var_fecha_documento, var_clave_moneda, var_tipo_Cambio, var_serie, "")
            'Cadena = " insert into tb_encabezado_cartera (VCHA_EMP_EMPRESA_ID,VCHA_UOR_UNIDAD_ID,VCHA_CAR_TIPO_DOCUMENTO, vcha_car_documento, Vcha_car_clase_id, inte_car_numero, char_car_afectacion, "
            'Cadena = Cadena + " vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_emo_numero, dtim_Car_fecha, VCHA_AGE_AGENTE_ID, Vcha_gac_grupo_actual_id, Vcha_gre_grupo_real_id, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, vcha_esb_establecimiento_ID, INTE_CAR_PLAZO, floa_car_porcentaje_iva, floa_Car_porcentaje_impuesto_1, floa_car_porcentaje_impuesto_2, floa_car_porcentaje_descuento_1, floa_car_porcentaje_descuento_2, floa_car_porcentaje_Descuento_3, floa_car_importe_total, floa_car_importe_iva,floa_car_importe_impuesto_1, floa_car_importe_impuesto_2, floa_car_importe_descuento_1, floa_car_importe_descuento_2, floa_car_importe_descuento_3, floa_car_subimporte, floa_car_importe_neto, vcha_car_importe_letra, Vcha_aud_usuario, Vcha_aud_maquina, vcha_aud_fecha, floa_Car_saldo, dtim_car_fecha_vencimiento, dtim_car_fecha_entrega, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, vcha_car_referencia) values "
            
            'Cadena = Cadena + "('" + rsaux4!vcha_emp_empresa_id + "', '', 'NC', 'DF', 'DF', " + CStr(CDbl(rsaux4!inte_mvs_numero)) + ", '-', '', '', 0, '" + CStr(rsaux4!DTIM_MVS_FECHA_DOCUMENTO) + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + var_cliente + "', '', 0, " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", "
            'Cadena = Cadena + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '" + CStr(var_fecha_documento) + "', 0, '" + CStr(var_fecha_documento) + "', '" + CStr(var_fecha_documento) + "', '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', '', '')"
            'rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_importe = 0
            var_subimporte = 0
            var_importe_iva = 0
         'End If
         rsaux4.MoveNext
   Wend
   rsaux4.Close
   End If
x = 0
If x = 0 Then
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   rsaux3.Open "select * from VW_TEMP_RELACION_COBRANZA_DF", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux3.EOF
         var_numero_nota = rsaux3!inte_mvs_numero
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + rsaux3!vcha_Rco_folio + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_iva = IIf(IsNull(rsaux!floa_rco_iva), 0, rsaux!floa_rco_iva)
            var_clave_moneda = IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id)
            var_tipo_Cambio = IIf(IsNull(rsaux!floa_rco_tipo_cambio), 1, rsaux!floa_rco_tipo_cambio)
         Else
            var_iva = 0
            var_clave_moneda = ""
            var_tipo_Cambio = 1
         End If
         If var_tipo_Cambio = 0 Then
            var_tipo_Cambio = 1
         End If
         var_fecha_documento = rsaux3!DTIM_MVS_FECHA_DOCUMENTO
         var_serie_cargo = rsaux3!vcha_mvs_serie_documento_abono
         var_documento = rsaux3!vcha_mvs_clave_documento_abono
         var_importe = (IIf(IsNull(rsaux3!floa_mvs_importe_neto), 0, rsaux3!floa_mvs_importe_neto))
         var_descuento = IIf(IsNull(rsaux3!floa_mva_descuento), 0, rsaux3!floa_mva_descuento)
         
         var_importe_total = var_importe - (var_importe / (1 + (var_descuento / 100)))
         var_factura = rsaux3!inte_mvs_numero_documento_abono
         var_descuento_otorgado = IIf(IsNull(rsaux3!floa_mva_descuento), 0, rsaux3!floa_mva_descuento) * 1
         var_descuento_aplicado = IIf(IsNull(rsaux3!floa_mva_descuento), 0, rsaux3!floa_mva_descuento) * 1
         'var_numero_nota = rsaux3!inte_mvs_numero
         var_numero_nota = CDbl(rsaux3!vcha_mvs_numero_documento)
         var_serie = rsaux3!vcha_mvs_serie
         If rsaux4.State = 1 Then
            rsaux4.Close
         End If
         rsaux4.Open "SELECT * FROM TB_RELACION_COBRANZA where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + rsaux3!vcha_Rco_folio + "' and inte_car_numero = " + CStr(var_factura), cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux4.EOF
               var_partida = rsaux4!inte_rco_partida
               rs.Open "update tb_relacion_cobranza set INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = " + Str(var_numero_nota) + ", DTIM_RCO_FECHA_DESCUENTO_FINANCIERO = '" + CStr(var_fecha_documento) + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + rsaux3!vcha_Rco_folio + "' and inte_car_numero = " + CStr(var_factura) + " and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
               rsaux4.MoveNext
         Wend
         var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(rsaux3!VCHA_EMP_EMPRESA_ID, var_serie_cargo, var_documento, rsaux3!inte_mvs_numero_documento_abono, var_serie, "DF", var_numero_nota, 0, (var_importe * var_tipo_Cambio))
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "Insert into TB_DETALLE_DESCUENTOS_FINANCIEROS (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, vcha_car_clase_id, inte_car_numero, vcha_ddf_concepto, floa_ddf_importe, inte_ddf_factura, floa_ddf_iva, floa_ddf_descuento_otorgado, floa_ddf_descuento_aplicado) values ('" + rsaux3!VCHA_EMP_EMPRESA_ID + "', 'DF', '" + var_serie + "','DF'," + Str(var_numero_nota) + ",'', " + Str((var_importe * var_tipo_Cambio)) + ", " + Str(var_factura) + ", " + CStr(var_iva) + ", " + CStr(var_descuento_otorgado) + ", " + CStr(var_descuento_aplicado) + " )", cnn, adOpenDynamic, adLockOptimistic
         var_subimporte = var_importe / (1 + (var_iva / 100))
         var_importe_iva = var_importe - var_subimporte
         var_numero_nota = var_numero_nota + 1
         var_importe = 0
         var_importe_iva = 0
         rsaux3.MoveNext
    Wend
  End If
  cnn.CommitTrans
   MsgBox "Termino", vbOKOnly, ""
End Sub

Private Sub Command3_Click()
Dim var_importe_neto_1 As Double
Dim var_importe_total_1 As Double
Dim var_subimporte_1 As Double
Dim var_importe_iva_1 As Double

Dim var_tipo_Cambio As Double
Dim var_importe_factura As Double
Dim var_importe_pago As Double
Dim var_importe_saldo_pago As Double
Dim var_importe_total As Double
Dim var_fecha_pago As Date
Dim var_fecha_factura As Date
Dim var_contador_pagos As Double
Dim var_contador_facturas As Double
Dim var_descuento_agente As Double
Dim var_descuento_sistema As Double
Dim var_saldo As Double
Dim si As Integer
Dim i, n As Integer
Dim var_importe As Double
Dim var_descuento As Double
Dim var_importe_descuento As Double
Dim var_moneda_local As Integer
Dim var_posible_tipo_cambio As Boolean
Dim var_numero_folio As Double
Dim var_serie_cargo As String
Dim var_importe_neto As Double
Dim var_subimporte As Double
Dim var_importe_iva As Double
Dim var_numero_nota_inicio As Integer
Dim var_k As Integer
Dim var_l As Integer
Dim var_numero_nota As Integer
Dim var_contador_notas As Integer
Dim var_referencia As String
Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
Set TB_ENCABEZADO_CARTERA_MIGRACION = New TB_ENCABEZADO_CARTERA_MIGRACION
   cnn.BeginTrans
   Dim var_fecha_documento As Date
   rsaux4.Open "select * from VW_TEMP_BONIFICACIONES_DEVOLUCIONES_SUMATORIA", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux4.EOF
          var_referencia = IIf(IsNull(rsaux4!vcha_mvs_numero_documento), "", rsaux4!vcha_mvs_numero_documento)
         var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
         var_serie = rsaux4!vcha_mvs_serie
         var_tipo_Cambio = 1
         var_tipo_Cambio = rsaux4!floa_MVS_TIPO_cambio
         If var_tipo_Cambio = 0 Then
            var_tipo_Cambio = 1
         End If
         var_fecha_documento = rsaux4!DTIM_MVS_FECHA_DOCUMENTO
         var_importe_neto = rsaux4!importe_neto
         var_subimporte = rsaux4!importe_neto
         var_iva = IIf(IsNull(rsaux4!FLOA_TPE_IVA), 0, rsaux4!FLOA_TPE_IVA)
         var_importe_iva = var_importe_neto - (var_importe_neto / (1 + (var_iva / 100)))
         var_subimporte = var_importe_neto - var_importe_iva
         
         Cadena = " insert into tb_encabezado_cartera (VCHA_EMP_EMPRESA_ID,VCHA_UOR_UNIDAD_ID,VCHA_CAR_TIPO_DOCUMENTO, vcha_car_documento, Vcha_car_clase_id, inte_car_numero, char_car_afectacion, "
         Cadena = Cadena + " vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_emo_numero, dtim_Car_fecha, VCHA_AGE_AGENTE_ID, Vcha_gac_grupo_actual_id, Vcha_gre_grupo_real_id, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, vcha_esb_establecimiento_ID, INTE_CAR_PLAZO, floa_car_porcentaje_iva, floa_Car_porcentaje_impuesto_1, floa_car_porcentaje_impuesto_2, floa_car_porcentaje_descuento_1, floa_car_porcentaje_descuento_2, floa_car_porcentaje_Descuento_3, floa_car_importe_total, floa_car_importe_iva,floa_car_importe_impuesto_1, floa_car_importe_impuesto_2, floa_car_importe_descuento_1, floa_car_importe_descuento_2, floa_car_importe_descuento_3, floa_car_subimporte, floa_car_importe_neto, vcha_car_importe_letra, Vcha_aud_usuario, Vcha_aud_maquina, vcha_aud_fecha, floa_Car_saldo, dtim_car_fecha_vencimiento, dtim_car_fecha_entrega, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, vcha_car_referencia) values "
         Cadena = Cadena + "('" + rsaux4!VCHA_EMP_EMPRESA_ID + "', '', 'NC', '" + rsaux4!vcha_mvs_clave_documento + "', '" + rsaux4!vcha_mvs_clave_documento + "', " + CStr(CDbl(rsaux4!inte_mvs_numero)) + ", '-', '', '', 0, '" + CStr(rsaux4!DTIM_MVS_FECHA_DOCUMENTO) + "', '" + IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID) + "', '" + IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID) + "', '" + IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id) + "', '" + IIf(IsNull(rsaux4!vcha_tit_titular_id), "", rsaux4!vcha_tit_titular_id) + "', '" + IIf(IsNull(rsaux4!vcha_cli_clave_id), "", rsaux4!vcha_cli_clave_id) + "', '', 0, " + CStr(IIf(IsNull(rsaux4!FLOA_TPE_IVA), 0, rsaux4!FLOA_TPE_IVA)) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_neto) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", "
         Cadena = Cadena + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', '" + CStr(var_fecha_documento) + "', 0, '" + CStr(var_fecha_documento) + "', '" + CStr(var_fecha_documento) + "', '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', '', '')"
         If rsaux3.State = 1 Then
            rsaux3.Close
         End If
         rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         
         'var_insertar = TB_ENCABEZADO_CARTERA_MIGRACION.Anadir(rsaux4!VCHA_EMP_EMPRESA_ID, "", "NC", rsaux4!vcha_mvs_clave_documento, rsaux4!vcha_mvs_clave_documento, CDbl(rsaux4!inte_mvs_numero), "-", "", "", 0, CStr(rsaux4!DTIM_MVS_FECHA_DOCUMENTO), IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID), IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_ACTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_ACTUAL_ID), IIf(IsNull(rsaux4!VCHA_GRE_GRUPO_REAL_ID), "", rsaux4!VCHA_GRE_GRUPO_REAL_ID), IIf(IsNull(rsaux4!VCHA_TIT_TITULAR_ID), "", rsaux4!VCHA_TIT_TITULAR_ID), IIf(IsNull(rsaux4!VCHA_CLI_CLAVE_ID), "", rsaux4!VCHA_CLI_CLAVE_ID), "", 0, IIf(IsNull(rsaux4!FLOA_TPE_IVA), 0, rsaux4!FLOA_TPE_IVA), 0, 0, 0, 0, 0, var_importe_neto, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, var_fecha_documento, 0, var_fecha_documento, var_fecha_documento, var_clave_moneda, var_tipo_Cambio, var_serie, "")
         rsaux4.MoveNext
   Wend
   rsaux4.Close
   rsaux4.Open "select * from VW_TEMP_BONIFICACIONES_DEVOLUCIONES", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux4.EOF
         var_iva = rsaux4!FLOA_TPE_IVA
         var_importe_neto = rsaux4!floa_mvs_importe_neto
         var_iva = rsaux4!FLOA_TPE_IVA
         var_tipo_Cambio = IIf(IsNull(rsaux4!floa_MVS_TIPO_cambio), 1, rsaux4!floa_MVS_TIPO_cambio)
         If var_tipo_Cambio = 0 Then
            var_tipo_Cambio = 1
         End If
         rs.Open "insert into tb_estado_cuenta (vcha_emp_empresa_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_ecu_numero_cargo, vcha_ecu_serie_abono, vcha_ecu_movimiento_abono, inte_ecu_numero_abono, floa_ecu_importe_Cargo, floa_ecu_importe_abono) values ('" + rsaux4!VCHA_EMP_EMPRESA_ID + "', '" + rsaux4!vcha_mvs_serie_documento_abono + "' ,'" + Trim(rsaux4!vcha_mvs_clave_documento_abono) + "', " + CStr(rsaux4!inte_mvs_numero_documento_abono) + ",'" + rsaux4!vcha_mvs_serie + "' ,'" + rsaux4!vcha_mvs_clave_documento + "'," + CStr(rsaux4!inte_mvs_numero) + ", 0, " + Str(var_importe_neto) + ")", cnn, adOpenDynamic, adLockOptimistic
         If rsaux4!vcha_mvs_clave_documento = "BO" Then
            rs.Open "insert into tb_detalle_bonificaciones (vcha_emp_empresa_id, vcha_car_documento,vcha_car_clase_id, vcha_ser_serie_id, inte_car_numero, inte_car_factura, floa_dbo_importe, floa_dbo_iva, char_dbo_estatus) values ('" + rsaux4!VCHA_EMP_EMPRESA_ID + "', 'BO', '','" + rsaux4!vcha_mvs_serie + "', " + CStr(rsaux4!inte_mvs_numero) + "," + CStr(rsaux4!inte_mvs_numero_documento_abono) + ", " + Str(rsaux4!floa_mvs_importe_neto) + "," + Str(IIf(IsNull(var_iva), 0, var_iva)) + ",'')"
         End If
         rsaux4.MoveNext
   Wend
cnn.CommitTrans
End Sub

Private Sub Command4_Click()
   Dim var_fecha_cobranza As Date
   Set TB_ENCABEZADO_CARTERA_MIGRACION = New TB_ENCABEZADO_CARTERA_MIGRACION
   Set TB_DEVOLUCIONES_ESTATUS = New TB_DEVOLUCIONES_ESTATUS
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   cnn.BeginTrans
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   rsaux4.Open "select count(*) from tb_relacion_cobranza WHERE CHAR_RCO_APLICADA <> '*'", cnn, adOpenDynamic, adLockOptimistic
   Label5 = CStr(rsaux4(0).Value)
   rsaux4.Close
   
   rsaux4.Open "select * from tb_relacion_cobranza WHERE CHAR_RCO_APLICADA <> '*'", cnn, adOpenDynamic, adLockOptimistic
   var_i = 0
   While Not rsaux4.EOF
         var_i = var_i + 1
         Label4 = CStr(var_i)
         Me.Refresh
         Label4.Refresh
         var_serie = rsaux4!VCHA_SER_SERIE_ID
         var_banco = "EFVO"
         var_empresa = rsaux4!VCHA_EMP_EMPRESA_ID
         var_fecha_cobranza = rsaux4!dtim_rco_fecha_relacion
         var_movimiento = 0
         var_numero_movimiento = 0
         var_almacen = ""
         var_partida = rsaux4!inte_rco_partida
         var_tipo_documento = IIf(IsNull(rsaux4!vcha_Car_documento), "", rsaux4!vcha_Car_documento)
         var_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
         var_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
         var_titular = IIf(IsNull(rsaux4!vcha_tit_titular_id), "", rsaux4!vcha_tit_titular_id)
         var_agente = IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID)
         var_cliente = rsaux4!vcha_cli_clave_id
         var_establecimiento = IIf(IsNull(rsaux4!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux4!vcha_ESB_ESTABLECIMIENTO_id)
         var_clave_moneda = IIf(IsNull(rsaux4!vcha_mon_moneda_id), "", rsaux4!vcha_mon_moneda_id)
         var_importe_total_cobranza = IIf(IsNull(rsaux4!floa_rco_importe), 0, rsaux4!floa_rco_importe)
         var_cheque = "EFVO"
         var_descuento_sistema = IIf(IsNull(rsaux4!floa_Car_descuento_aplicado), 0, rsaux4!floa_Car_descuento_aplicado)
         var_descuento_agente = IIf(IsNull(rsaux4!FLOA_RCO_DESCUENTO_OTORGADO), 0, rsaux4!FLOA_RCO_DESCUENTO_OTORGADO)
         var_descuento_aplicar = 0
         var_porcentaje_iva = IIf(IsNull(rsaux4!floa_rco_iva), 0, rsaux4!floa_rco_iva)
         var_porcentaje_impuesto_2 = IIf(IsNull(rsaux4!floa_rco_impuesto_2), 0, rsaux4!floa_rco_impuesto_2)
         var_porcentaje_impuesto_3 = IIf(IsNull(rsaux4!floa_rco_impuesto_3), 0, rsaux4!floa_rco_impuesto_3)
         var_tipo_Cambio = IIf(IsNull(rsaux4!floa_rco_tipo_cambio), 1, rsaux4!floa_rco_tipo_cambio)
         If var_tipo_Cambio = 0 Then
            var_tipo_Cambio = 1
         End If
         If var_descuento_agente < var_descuento_sistema Then
            var_descuento_aplicar = var_descuento_agente
         End If
         If var_descuento_sistema < var_descuento_agente Then
            var_descuento_aplicar = var_descuento_sistema
         End If
         If var_descuento_sistema = var_descuento_agente Then
            var_descuento_aplicar = var_descuento_sistema
         End If
         var_insertar = False
         var_importe_cobranza = (var_importe_total_cobranza * 100) / (100 - var_descuento_aplicar)
         If var_importe_saldo < var_importe_cobranza Then
            If var_importe_saldo > 0 Then
               var_descuento_aplicar = 100 - ((var_importe_total_cobranza * 100) / var_importe_saldo)
            Else
               var_descuento_aplicar = 0
            End If
            rsaux2.Open "update tb_relacion_cobranza set floa_rco_descuento_aplicar = " + CStr(var_descuento_aplicar) + ", dtim_rco_fecha_aplicacion =  '" + CStr(var_fecha_cobranza) + "'  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_cheque = '" + var_cheque + "' and vcha_cli_clave_id = '" + var_cliente + "' and vcha_ban_banco_id = '" + var_banco + "' and inte_car_numero = " + CStr(IIf(IsNull(rsaux4!inte_Car_numero), 0, rsaux4!inte_Car_numero)) + " and vcha_rco_folio = '" + IIf(IsNull(rsaux4!vcha_Rco_folio), "", rsaux4!vcha_Rco_folio) + "' and vcha_car_documento = '" + Trim(var_tipo_documento) + "' and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
         End If
         var_importe_total_cobranza = var_importe_total_cobranza
         var_importe_total = (var_importe_total_cobranza) * (1 + (var_descuento_aplicar / 100))
         var_subimporte = var_importe_total / (1 + (var_descuento_aplicar / 100))
         var_importe_descuento_1 = var_importe_total - var_subimporte
         var_importe_descuento_2 = 0
         var_importe_descuento_3 = 0
         var_importe_iva = var_importe_total_cobranza - (var_importe_total_cobranza) / (1 + (var_porcentaje_iva / 100))
         If var_porcentaje_impuesto_2 > 0 Then
            var_importe_impuesto_2 = (var_importe_total_cobranza - var_importe_iva) / (var_importe_total_cobranza - var_importe_iva) / (1 + (var_porcentaje_impuesto_2 / 100))
         Else
            var_importe_impuesto_2 = 0
         End If
         If var_porcentaje_impuesto_3 > 0 Then
            var_importe_impuesto_3 = (var_importe_total_cobranza - var_importe - iva_var_impuesto_2) / (var_importe_total_cobranza - var_importe_iva - var_impuesto_2) / (1 + (var_porcentaje_impuesto_3 / 100))
         Else
            var_importe_impuesto_3 = 0
         End If
         rsaux3.Open "select maximo_numero from vw_maximo_numero_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_car_tipo_documento = 'PA' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux3.EOF Then
            var_numero_folio = 0
         Else
            var_numero_folio = IIf(IsNull(rsaux3!maximo_numero), 0, rsaux3!maximo_numero)
         End If
         rsaux3.Close
         var_numero_folio = var_numero_folio + 1
         rsaux3.Open "update tb_relacion_cobranza set char_rco_aplicada = '*', FLOA_RCO_TIPO_CAMBIO = " + Str(var_tipo_Cambio) + ", INTE_RCO_PAGO = " + Str(var_numero_folio) + ", dtim_rco_fecha_aplicacion =  '" + CStr(var_fecha_cobranza) + "'   where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_cheque = 'EFVO' and vcha_cli_clave_id = '" + var_cliente + "' and vcha_ban_banco_id = 'EFVO' and vcha_rco_folio = '" + IIf(IsNull(rsaux4!vcha_Rco_folio), "", rsaux4!vcha_Rco_folio) + "' and inte_Car_numero = " + CStr(rsaux4!inte_Car_numero) + " and vcha_car_documento = '" + var_tipo_documento + "' and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
         var_importe_sin_impuesto = var_importe_total_cobranza - (var_importe_iva + var_importe_descuento_2 + var_importe_descuento_3)
         
         var_insertar = TB_ENCABEZADO_CARTERA_MIGRACION.Anadir(var_empresa, var_unidad_organizacional, "PA", "PA", "PA", var_numero_folio, "-", "", "", 0, CStr(var_fecha_cobranza), var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_porcentaje_iva, var_porcentaje_impuesto_2, var_porcentaje_impuesto_3, var_descuento_aplicar, 0, 0, var_importe_total, var_importe_iva, var_importe_impuesto_2, var_importe_impuesto_3, var_importe_descuento_1, 0, 0, var_importe_sin_impuesto, var_importe_total_cobranza, "", var_clave_usuario_global, "", var_fecha_cobranza, 0, var_fecha_cobranza, var_fecha_cobranza, var_clave_moneda, var_tipo_Cambio, var_serie, "")
         var_insertar = False
         var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, var_tipo_documento, rsaux4!inte_Car_numero, var_serie, "PA", var_numero_folio, 0, var_importe_total_cobranza)
         If var_tipo_documento = "FA" Then
            rsaux3.Open "select * from VW_DETALLE_FACTURACION_LINEAS WHERE VCHA_EMP_EMPRESA_ID =  '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' and inte_Car_numero = " + CStr(rsaux4!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux3.EOF
                  var_fecha_factura = CDate(Format(CStr(rsaux3!dtim_Car_fecha), "short date"))
                  var_cadena = "INSERT INTO TB_COMISIONES_APLICADAS ([VCHA_EMP_EMPRESA_ID], [VCHA_AGE_AGENTE_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_SER_SERIE_ID], [INTE_CAR_NUMERO], [DTIM_CAR_FECHA], [FLOA_CAP_IMPORTE_FACTURA], [VCHA_RCO_FOLIO], [DTIM_CAP_FECHA_PAGO], [VCHA_LIN_LINEA_ID], [FLOA_CAP_IMPORTE_PARTICIPACION], [FLOA_CAP_PORCENTAJE_PARTICIPACION], [FLOA_COM_PORCENTAJE], [FLOA_CAP_IMPORTE_COMISION], [VCHA_BAN_BANCO_ID] , [VCHA_RCO_CHEQUE], [FLOA_CAP_IMPORTE_PAGO], [VCHA_CLI_CLAVE_ID])"
                  var_cadena = var_cadena + "Values ('" + var_empresa + "', '" + var_agente + "', 'FA', '" + var_serie + "', " + CStr(rsaux4!inte_Car_numero) + ", '" + Str(Day(var_fecha_factura)) + "/" + Str(Month(var_fecha_factura)) + "/" + Str(Year(var_fecha_factura)) + "', " + CStr(rsaux3!floa_Car_importe_neto / rsaux3!floa_car_tipo_cambio) + ", '" + rsaux4!vcha_Rco_folio + "', '" + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date)) + "', '" + rsaux3!vcha_lin_linea_id + "', " + CStr(rsaux3!Importe / rsaux3!floa_car_tipo_cambio) + ", 0, 0, 0,'EFVO' ,'EFVO', " + CStr(var_importe_sin_impuesto) + ", '" + var_cliente + "')"
                  rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux3.MoveNext
            Wend
            rsaux3.Close
         End If
         rsaux4.MoveNext
   Wend
   cnn.CommitTrans
 End Sub

Private Sub Command5_Click()
        Set TB_GRUPOSACTUALES = New TB_GRUPOSACTUALES
        Set TB_GRUPOSREALES = New TB_GRUPOSREALES
        Set TB_TITULARES = New TB_TITULARES
        Set TB_CLIENTES = New TB_CLIENTES
        Set TB_ESTABLECIMIENTOS = New TB_ESTABLECIMIENTOS
   rs.Open "select * from agente56", var_tabla, adOpenDynamic, adLockOptimistic
   Dim var_nombre As String
   While Not rs.EOF
        txt_cliente = ""
        var_modifica_regsitro_establecimientos = False
        var_modifica_registro_cliente = False
        var_modifica_registro_titular = False
        var_modifica_registro_gr = False
        var_modifica_registro_ga = False
        var_nombre = IIf(IsNull(rs!razonsocia), "", rs!razonsocia)
        cnn.BeginTrans
        ok = TB_GRUPOSACTUALES.Anadir("", var_nombre, 0, 0)
        ok = TB_GRUPOSREALES.Anadir(var_grupo_actual_regreso, "x", Trim(var_nombre), txt_descuento_1, txt_descuento_2, txt_descuento_3, txt_prioridad)
        ok = TB_TITULARES.Anadir(var_grupo_real_regreso, "", var_nombre, "", "", "", "", "", IIf(IsNull(rs!DIRECCION), "", rs!DIRECCION), "", txt_limite_credito, IIf(IsNull(rs!CODIGOPOST), 0, rs!CODIGOPOST))
        ok = TB_CLIENTES.Anadir("", var_nombre, var_nombre, IIf(IsNull(rs!fechaalta), Date, rs!fechaalta), "00075", "", "", IIf(IsNull(rs!rfc), "", rs!rfc), "", 0, "", "", "", "", 0, 1, var_titular_regreso, "", "", "", "", "", "", IIf(IsNull(rs!DIRECCION), "", rs!DIRECCION), IIf(IsNull(rs!CODIGOPOST), "", rs!CODIGOPOST), "", 1, 1, rs!cvecliente)
        ok = TB_ESTABLECIMIENTOS.Anadir(var_titular_regreso, "", var_nombre, "", "", "", "", IIf(IsNull(rs!DIRECCION), "", rs!DIRECCION), "", "", "", IIf(IsNull(rs!CODIGOPOST), "", rs!CODIGOPOST))
        rsaux.Open "insert into TB_DETALLE_ESTABLECIMIENTOS (vcha_esb_establecimiento_id, vcha_cli_clave_id) values( '" + var_establecimiento_regreso + "', '" + var_cliente_regreso + "')", cnn, adOpenDynamic, adLockOptimistic
        cnn.CommitTrans
        rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command6_Click()
   cnn.BeginTrans
   rs.Open "update tb_temp_movsclie set vcha_rco_folio = vcha_mvs_numero_documento", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select * from TB_TEMP_MOVSCLIE where VCHA_MVS_CLAVE_DOCUMENTO = 'PA'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_numero_factura = rs!inte_mvs_numero_documento_abono
         var_serie = rs!vcha_mvs_serie_documento_abono
         If Not IsNull(rs!vcha_Rco_folio) Then
            rsaux.Open "update tb_temp_movsclie set vcha_rco_folio = '" + rs!vcha_Rco_folio + "' where INTE_MVS_NUMERO_DOCUMENTO_ABONO = " + Str(var_numero_factura) + " and VCHA_MVS_SERIE_DOCUMENTO_ABONO = '" + var_serie + "' and VCHA_MVS_CLAVE_DOCUMENTO  = 'DF'", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Open "update tb_temp_movsclie set vcha_rco_folio = '" + rs!vcha_Rco_folio + "' where INTE_MVS_NUMERO_DOCUMENTO_ABONO = " + Str(var_numero_factura) + " and VCHA_MVS_SERIE_DOCUMENTO_ABONO = '" + var_serie + "' and VCHA_MVS_CLAVE_DOCUMENTO  = 'BF'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.MoveNext
   Wend
   rs.Close
   
   'rs.Open "select distinct vcha_emp_empresa_id, vcha_mva_clave_abono, vcha_mva_numero_abono, vcha_mva_numero_documento from tb_temp_movsabon", cnn, adOpenDynamic, adLockOptimistic
   'While Not rs.EOF
   '      rsaux.Open "update tb_temp_movsclie set vcha_rco_folio = '" + rs!vcha_mva_numero_documento + "' where vcha_emp_empresa_id = '" + rs!vcha_emp_empresa_id + "' and vcha_mvs_numero_documento = '" + rs!vcha_mva_numero_abono + "' and vcha_mvs_clave_documento = '" + rs!vcha_mva_clave_abono + "' and VCHA_MVS_CLAVE_DOCUMENTO  = 'DF'", cnn, adOpenDynamic, adLockOptimistic
   '      rsaux.Open "update tb_temp_movsclie set vcha_rco_folio = '" + rs!vcha_mva_numero_documento + "' where vcha_emp_empresa_id = '" + rs!vcha_emp_empresa_id + "' and vcha_mvs_numero_documento = '" + rs!vcha_mva_numero_abono + "' and vcha_mvs_clave_documento = '" + rs!vcha_mva_clave_abono + "' and VCHA_MVS_CLAVE_DOCUMENTO  = 'BF'", cnn, adOpenDynamic, adLockOptimistic
   '      rs.MoveNext
   'Wend
   'rs.Close
   cnn.CommitTrans
End Sub

Private Sub Command7_Click()
   Dim var_tolerancia_saldo As Integer
   Dim var_descuento_otorgado As Double
   Dim var_descuento_aplicado As Double
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_contador_lineas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim si, i, n As Integer
   Dim var_saldo As Double
   Dim var_serie_cargo As String
   Dim var_numero_nota_inicio As Integer
   Dim var_factura As Double
   Dim var_k As Integer
   Dim var_descuentos As Integer
   Dim var_desc_otorgado_str As String
   Dim var_desc_apilcado_str As String
   Dim var_documento As String
   Dim var_fecha_documento As Date
   Set TB_ENCABEZADO_CARTERA_MIGRACION = New TB_ENCABEZADO_CARTERA_MIGRACION
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   var_tolerancia_saldo = 6
   x = 1
   cnn.BeginTrans
   If x = 1 Then
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   rsaux4.Open "SELECT * FROM VW_TEMP_RELACION_COBRANZA_SUMATORIA_BF", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux4.EOF
         var_almacen = ""
         var_grupo_actual = rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID
         var_grupo_real = rsaux4!vcha_gre_grupo_real_id
         var_titular = rsaux4!vcha_tit_titular_id
         var_agente = IIf(IsNull(rsaux4!VCHA_AGE_AGENTE_ID), "", rsaux4!VCHA_AGE_AGENTE_ID)
         var_cliente = rsaux4!vcha_cli_clave_id
         var_fecha_documento = rsaux4!DTIM_MVS_FECHA_DOCUMENTO
         var_establecimiento = rsaux4!vcha_ESB_ESTABLECIMIENTO_id
         If rsaux2.State = 1 Then
            rsaux2.Close
         End If
         rsaux2.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + rsaux4!vcha_Rco_folio + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            var_iva = IIf(IsNull(rsaux2!floa_rco_iva), 0, rsaux2!floa_rco_iva)
            var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
            var_tipo_Cambio = IIf(IsNull(rsaux2!floa_rco_tipo_cambio), 1, rsaux2!floa_rco_tipo_cambio)
         Else
            var_iva = 0
            var_clave_moneda = ""
            var_tipo_Cambio = 1
         End If
         var_serie = rsaux4!vcha_mvs_serie
         If var_tipo_Cambio = 0 Then
            var_tipo_Cambio = 1
         End If
         var_numero_nota = CDbl(rsaux4!vcha_mvs_numero_documento)
         var_importe = rsaux4!importe_neto
         var_subimporte = var_importe / (1 + (var_iva / 100))
         var_importe_iva = var_importe - var_subimporte
         Cadena = " insert into tb_encabezado_cartera (VCHA_EMP_EMPRESA_ID,VCHA_UOR_UNIDAD_ID,VCHA_CAR_TIPO_DOCUMENTO, vcha_car_documento, Vcha_car_clase_id, inte_car_numero, char_car_afectacion, "
         Cadena = Cadena + " vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_emo_numero, dtim_Car_fecha, VCHA_AGE_AGENTE_ID, Vcha_gac_grupo_actual_id, Vcha_gre_grupo_real_id, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, vcha_esb_establecimiento_ID, INTE_CAR_PLAZO, floa_car_porcentaje_iva, floa_Car_porcentaje_impuesto_1, floa_car_porcentaje_impuesto_2, floa_car_porcentaje_descuento_1, floa_car_porcentaje_descuento_2, floa_car_porcentaje_Descuento_3, floa_car_importe_total, floa_car_importe_iva,floa_car_importe_impuesto_1, floa_car_importe_impuesto_2, floa_car_importe_descuento_1, floa_car_importe_descuento_2, floa_car_importe_descuento_3, floa_car_subimporte, floa_car_importe_neto, vcha_car_importe_letra, Vcha_aud_usuario, Vcha_aud_maquina, vcha_aud_fecha, floa_Car_saldo, dtim_car_fecha_vencimiento, dtim_car_fecha_entrega, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, vcha_car_referencia) values "
                                                                         Cadena = Cadena + "('" + rsaux4!VCHA_EMP_EMPRESA_ID + "', '', 'NC', 'BF', 'BF', " + CStr(var_numero_nota) + ", '-', '', '', 0, '" + CStr(var_fecha_documento) + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + var_cliente + "', '" + var_establecimiento + "', 0, " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe * var_tipo_Cambio) + ", " + CStr(var_importe_iva * var_tipo_Cambio) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte * var_tipo_Cambio) + ", " + CStr(var_importe * var_tipo_Cambio) + ", '', '" + var_clave_usuario_global + "', '', '" + CStr(var_fecha_documento) + "', 0, '" + CStr(var_fecha_documento) + "', '" + CStr(var_fecha_documento) + "', '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "','I','')"
         If rsaux3.State = 1 Then
            rsaux3.Close
         End If
         rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         'var_insertar = TB_ENCABEZADO_CARTERA_MIGRACION.Anadir(rsaux4!vcha_emp_empresa_id, "08", "NC", "BF", "BF", var_numero_nota, "-", "", "", 0, CStr(var_fecha_documento), var_agente, var_grupo_actual, var_grupo_real, var_titular, var_cliente, var_establecimiento, 0, var_iva, 0, 0, 0, 0, 0, var_importe * var_tipo_Cambio, var_importe_iva * var_tipo_Cambio, 0, 0, 0, 0, 0, var_subimporte * var_tipo_Cambio, var_importe * var_tipo_Cambio, "", var_clave_usuario_global, "", var_fecha_documento, 0, var_fecha_documento, var_fecha_documento, var_clave_moneda, var_tipo_Cambio, var_serie, "")
         var_importe = 0
         var_subimporte = 0
         var_importe_iva = 0
         rsaux4.MoveNext
   Wend
   rsaux4.Close
   End If
x = 0
If x = 0 Then
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   rsaux3.Open "select * from VW_TEMP_RELACION_COBRANZA_BF", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux3.EOF
         var_numero_nota = rsaux3!inte_mvs_numero
         var_fecha_documento = rsaux3!DTIM_MVS_FECHA_DOCUMENTO
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + rsaux3!vcha_Rco_folio + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_iva = IIf(IsNull(rsaux!floa_rco_iva), 0, rsaux!floa_rco_iva)
            var_clave_moneda = IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id)
            var_tipo_Cambio = IIf(IsNull(rsaux!floa_rco_tipo_cambio), 1, rsaux!floa_rco_tipo_cambio)
         Else
            var_iva = 0
            var_clave_moneda = ""
            var_tipo_Cambio = 1
         End If
         If var_tipo_Cambio = 0 Then
            var_tipo_Cambio = 1
         End If
         var_serie_cargo = rsaux3!vcha_mvs_serie_documento_abono
         var_documento = rsaux3!vcha_mvs_clave_documento_abono
         var_importe = (IIf(IsNull(rsaux3!floa_mvs_importe_neto), 0, rsaux3!floa_mvs_importe_neto))
         var_descuento = IIf(IsNull(rsaux3!floa_mva_descuento), "", rsaux3!floa_mva_descuento)
         var_importe_total = var_importe - (var_importe / (1 + (var_descuento / 100)))
         var_factura = rsaux3!inte_mvs_numero_documento_abono
         var_descuento_otorgado = IIf(IsNull(rsaux3!floa_mva_descuento), 0, rsaux3!floa_mva_descuento) * 1
         var_descuento_aplicado = IIf(IsNull(rsaux3!floa_mva_descuento), 0, rsaux3!floa_mva_descuento) * 1
         var_numero_nota = CDbl(rsaux3!vcha_mvs_numero_documento)
         var_serie = rsaux3!vcha_mvs_serie
         
         rs.Open "update tb_relacion_cobranza set inte_rco_numero_bonificacion_financiera = " + Str(var_numero_nota) + ", dtim_rco_fecha_bonificacion_financiera = '" + CStr(var_fecha_documento) + "', FLOA_RCO_BONIFICACION_FINANCIERA = " + CStr(rsaux3!floa_mva_descuento) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + rsaux3!vcha_Rco_folio + "' and inte_car_numero = " + CStr(var_factura), cnn, adOpenDynamic, adLockOptimistic
         var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(rsaux3!VCHA_EMP_EMPRESA_ID, var_serie_cargo, var_documento, rsaux3!inte_mvs_numero_documento_abono, var_serie, "BF", var_numero_nota, 0, (var_importe * var_tipo_Cambio))
         If rsaux.State = 1 Then
            rsaux.Close
         End If

         'rsaux.Open "Insert into TB_DETALLE_DESCUENTOS_FINANCIEROS (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, vcha_car_clase_id, inte_car_numero, vcha_ddf_concepto, floa_ddf_importe, inte_ddf_factura, floa_ddf_iva, floa_ddf_descuento_otorgado, floa_ddf_descuento_aplicado) values ('" + rsaux3!vcha_emp_empresa_id + "', 'DF', '" + var_serie + "','DF'," + Str(var_numero_nota) + ",'', " + Str((var_importe * var_tipo_cambio)) + ", " + Str(var_factura) + ", " + CStr(var_iva) + ", " + CStr(var_descuento_otorgado) + ", " + CStr(var_descuento_aplicado) + " )", cnn, adOpenDynamic, adLockOptimistic
         rsaux.Open "insert into tb_detalle_bonificacion_financiera (vcha_emp_empresa_id, vcha_car_documento, vcha_ser_serie_id, inte_Car_numero, inte_dbf_factura, floa_dbf_porcentaje, floa_dbf_importe, floa_dbf_iva) values ('" + rsaux3!VCHA_EMP_EMPRESA_ID + "', 'BF', '" + var_serie + "', " + CStr(var_numero_nota) + ", " + Str(var_factura) + ", " + CStr(rsaux3!floa_mva_descuento) + ", " + CStr((var_importe * var_tipo_Cambio)) + ", " + CStr(var_iva) + ")"
         var_subimporte = var_importe / (1 + (var_iva / 100))
         var_importe_iva = var_importe - var_subimporte
         var_numero_nota = var_numero_nota + 1
         var_importe = 0
         var_importe_iva = 0
         rsaux3.MoveNext
    Wend
  End If
  cnn.CommitTrans
End Sub

Private Sub Command8_Click()
   rsaux2.Open "update tb_temp_movsclie set vcha_cli_clave_id = ''", cnn, adOpenDynamic, adLockOptimistic
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select distinct vcha_cli_clave_anterior_id from tb_temp_movsclie", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select * from tb_clientes where vcha_cli_clave_anterior_id = '" + rs!VCHA_CLI_CLAVE_ANTERIOR_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux2.Open "update tb_temp_movsclie set vcha_cli_clave_id = '" + rsaux!vcha_cli_clave_id + "' where vcha_cli_clave_anterior_id = '" + rs!VCHA_CLI_CLAVE_ANTERIOR_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command9_Click()
    rs.Open "SELECT count(*) as veces FROM TB_ENCABEZADO_CARTERA where vcha_gac_grupo_actual_id is null or vcha_gac_grupo_actual_id = ''", cnn, adOpenDynamic, adLockOptimistic
    Label4 = rs!veces
    rs.Close
    var_i = 0
    rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA  where (vcha_age_agente_id is null or vcha_age_agente_id = '') or (vcha_gac_grupo_actual_id is null or vcha_gac_grupo_actual_id = '')", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
          If Not rsaux.EOF Then
             If rsaux2.State = 1 Then
                rsaux2.Close
             End If
             rsaux2.Open "UPDATE TB_ENCABEZADO_CARTERA SET char_car_estatus = 'I', vcha_age_agente_id = '" + IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID) + "', VCHA_GAC_GRUPO_ACTUAL_ID = '" + IIf(IsNull(rsaux!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux!VCHA_GAC_GRUPO_aCTUAL_ID) + "', VCHA_GRE_GRUPO_REAL_ID = '" + IIf(IsNull(rsaux!vcha_gre_grupo_real_id), "", rsaux!vcha_gre_grupo_real_id) + "', VCHA_TIT_TITULAR_ID = '" + IIf(IsNull(rsaux!vcha_tit_titular_id), "", rsaux!vcha_tit_titular_id) + "', VCHA_MON_MONEDA_ID = '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id) + "' WHERE VCHA_CAR_CLASE_ID = '" + rs!vcha_Car_clase_id + "' AND INTE_CAR_NUMERO = " + CStr(rs!inte_Car_numero) + " AND VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "'", cnn, adOpenDynamic, adLockOptimistic
          Else
             If rsaux2.State = 1 Then
                rsaux2.Close
             End If
             rsaux2.Open "UPDATE TB_ENCABEZADO_CARTERA SET char_car_estatus = 'I', vcha_gac_grupo_actual_id = '', vcha_gre_grupo_real_id = '',VCHA_TIT_TITULAR_ID = '' WHERE VCHA_CAR_CLASE_ID = '" + rs!vcha_Car_clase_id + "' AND INTE_CAR_NUMERO = " + CStr(rs!inte_Car_numero) + " AND VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND VCHA_CAR_TIPO_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "'", cnn, adOpenDynamic, adLockOptimistic
          End If
          rsaux.Close
          rs.MoveNext
          var_i = var_i + 1
          Label5 = CStr(var_i)
          Me.Refresh
    Wend
    rs.Close
    
    
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Dim var_ruta As String
   var_ruta = App.Path
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
End Sub

