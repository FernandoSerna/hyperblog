VERSION 5.00
Begin VB.Form frmmigracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migracion de Informacion"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8475
   Begin VB.CommandButton cmd_cartera 
      Caption         =   "Información de Cartera"
      Height          =   450
      Left            =   390
      TabIndex        =   12
      Top             =   3840
      Width           =   7920
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Enabled         =   0   'False
      Height          =   450
      Left            =   375
      TabIndex        =   11
      Top             =   3210
      Width           =   7920
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Actualizar "
      Enabled         =   0   'False
      Height          =   450
      Left            =   375
      TabIndex        =   10
      Top             =   2640
      Width           =   7920
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ejecutar migracion de información de clientes"
      Enabled         =   0   'False
      Height          =   450
      Left            =   420
      TabIndex        =   9
      Top             =   255
      Width           =   7920
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Relación entre los establecimientos y clientes"
      Enabled         =   0   'False
      Height          =   450
      Left            =   5880
      TabIndex        =   8
      Top             =   1965
      Width           =   2460
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Releciones entre los grupos"
      Enabled         =   0   'False
      Height          =   450
      Left            =   5880
      TabIndex        =   7
      Top             =   1515
      Width           =   2460
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tipos de clientes"
      Enabled         =   0   'False
      Height          =   450
      Left            =   5880
      TabIndex        =   6
      Top             =   1065
      Width           =   2460
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Grupos Actuales"
      Enabled         =   0   'False
      Height          =   450
      Left            =   3135
      TabIndex        =   5
      Top             =   1965
      Width           =   2460
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grupos Reales"
      Enabled         =   0   'False
      Height          =   450
      Left            =   3135
      TabIndex        =   4
      Top             =   1515
      Width           =   2460
   End
   Begin VB.CommandButton clientes 
      Caption         =   "clientes"
      Enabled         =   0   'False
      Height          =   450
      Left            =   3135
      TabIndex        =   3
      Top             =   1065
      Width           =   2460
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Establecimientos"
      Enabled         =   0   'False
      Height          =   450
      Left            =   405
      TabIndex        =   2
      Top             =   1965
      Width           =   2460
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Titulares"
      Enabled         =   0   'False
      Height          =   450
      Left            =   405
      TabIndex        =   1
      Top             =   1515
      Width           =   2460
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agentes"
      Enabled         =   0   'False
      Height          =   450
      Left            =   405
      TabIndex        =   0
      Top             =   1065
      Width           =   2460
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   300
      Left            =   4320
      TabIndex        =   16
      Top             =   4440
      Width           =   1860
   End
   Begin VB.Label Label3 
      Caption         =   "de"
      Height          =   300
      Left            =   3285
      TabIndex        =   15
      Top             =   4440
      Width           =   1860
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   300
      Left            =   1725
      TabIndex        =   14
      Top             =   4440
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   300
      Left            =   465
      TabIndex        =   13
      Top             =   4440
      Width           =   1860
   End
End
Attribute VB_Name = "frmmigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection

Private Sub clientes_Click()
   Dim fecha As Date
   Cadena = "select cvecliente,razonsocia,fechaalta,cveagente,rfc,cvetitular,tipoclient from clientes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      fecha = rsaux!fechaalta
      If fecha = "12:00:00 a.m." Then
         fecha = Date
      End If
      rs.Open "insert into tb_clientes (VCHA_CLI_CLAVE_ID,VCHA_CLI_NOMBRE,DTIM_CLI_FECHA_CAPTURA,VCHA_AGE_AGENTE_ID,VCHA_CLI_RFC,VCHA_TIT_TITULAR_ID,CHAR_PRI_PRIORIDAD_ID) values ('" + Trim(rsaux!cvecliente) + "','" + Trim(rsaux!razonsocia) + "','" + CStr(fecha) + "','" + rsaux!cveagente + "','" + Trim(rsaux!rfc) + "','" + Trim(rsaux!cvetitular) + "','" + Trim(rsaux!tipoclient) + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub cmd_cartera_Click()
   Dim var_numero As String
   Dim var_serie As String
   Dim var_numero_serie As String
   Dim var_numero_documento As Double
   Dim var_i As Integer, var_j As Integer
   Dim var_documento As String
   Dim var_tipo_documento As String
   Dim var_clase_documento As String
   Dim var_z As Integer
   Dim var_iva As Double, var_porcentaje_iva As Double, var_porcentaje_descuento_1 As Double, var_porcentaje_descuento_2 As Double, var_descuento_1 As Double, var_descuento_2 As Double
   rs.Open "delete from TB_TEM_MIGRACION_CARTERA", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT * FROM movsclie", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_numero = IIf(IsNull(rs!numdocumen), "0", rs!numdocumen)
         If IsNumeric(var_numero) Then
            var_serie = "X"
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
         End If
         Cadena = "INSERT INTO TB_TEM_MIGRACION_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_CLI_CLAVE_ID, DTIM_TEM_FECHA_CAPTURA, DTIM_TEM_FECHA_DOCUMENTO, CHAR_TEM_TIPO, VCHA_TEM_CLAVE_DOCUMENTO, VCHA_TEM_SERIE_DOCUMENTO, INTE_TEM_MUMERO_DOCUMENTO, VCHA_AGE_AGENTE_ANTERIOR_ID, VCHA_AGE_AGENTE_ID, VCHA_ZON_ZONA_ANTERIO_ID, VCHA_ZON_ZONA_ID, FLOA_TEM_COMISION, FLOA_TEM_IMPORTE_NETO, DTIM_TEM_FECHA_VENICIMIENTO, VCHA_TEM_REFERENCIA, FLOA_TEM_SALDO_DOCUMENTO)"
         Cadena = Cadena + " values ('" + rs!cveempresa + "', '" + rs!cvecliente + "', '', '" + CStr(rs!fechacaptu) + "', '" + CStr(rs!fechadocum) + "', '" + rs!tipo + "', '" + rs!cvedocumen + "', '" + var_serie + "', " + CStr(var_numero_documento) + ", '" + rs!cveagente + "', '', '" + rs!cvezona + "', '', " + CStr(rs!comision) + ", " + CStr(rs!importenet) + ", '" + CStr(rs!fechavenci) + "', '" + rs!referencia + "', " + CStr(rs!saldodocum) + ") "
         rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "select distinct vcha_cli_clave_anterior_id from TB_TEM_MIGRACION_CARTERA", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select vcha_cli_clave_id from tb_clientes where vcha_cli_clave_anterior_id = '" + rs!VCHA_CLI_CLAVE_ANTERIOR_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         var_clave_cliente = rsaux!vcha_cli_clave_id
         rsaux.Close
         rsaux.Open "select distinct vcha_age_agente_id, floa_tpe_iva, vcha_mon_moneda_id, inte_mon_moneda_local from vw_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         var_clave_agente = IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID)
         var_tipo_cambio = 0
         If var_moneda_local = 1 Then
            var_tipo_cambio = 1
         End If
         rsaux.Close
         rsaux2.Open "update TB_TEM_MIGRACION_CARTERA set vcha_cli_clave_id = '" + var_clave_cliente + "', vcha_age_agente_id = '" + var_clave_agente + "' where vcha_cli_clave_anterior_id = '" + rs!VCHA_CLI_CLAVE_ANTERIOR_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "SELECT cveempresa,numfactura ,descuento, dctoptopag, iva, comision FROM ARMANDO1", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_descuento_1 = 0
         var_descuento_2 = 0
         var_descuento_3 = 0
         var_porcentaje_iva = 0
         var_tipo_cambio = 1
         var_descuento_1 = IIf(IsNull(rs!descuento), 0, rs!descuento)
         var_descuento_2 = IIf(IsNull(rs!dctoptopag), 0, rs!dctoptopag)
         var_descuento_3 = 0
         var_porcentaje_iva = IIf(IsNull(rs!iva), 0, rs!iva)
         var_tipo_cambio = IIf(IsNull(rs!comision), 1, rs!comision)
         
         var_numero = IIf(IsNull(rs!numfactura), "0", rs!numfactura)
         If IsNumeric(var_numero) Then
            var_serie = "X"
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
         End If
         Cadena = "update TB_TEM_MIGRACION_CARTERA set FLOA_TEM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_TEM_dESCUENTO_2 =  " + CStr(var_descuento_2) + ", FLOA_TEM_PORCENTAJE_IVA = " + CStr(var_porcentaje_iva) + ", FLOA_TEM_TIPO_CAMBIO = " + CStr(var_tipo_cambio) + " WHERE VCHA_TEM_CLAVE_DOCUMENTO = 'FA' AND VCHA_TEM_SERIE_DOCUMENTO = '" + var_serie + "' AND INTE_TEM_MUMERO_DOCUMENTO = " + CStr(var_numero_documento)
         rsaux.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
'de aqui
   rs.Close
   rs.Open "SELECT cveempresa, numfactura, cvecliente, descuento, descuento2, dctoptopag, comision, iva FROM armando1", var_tabla, adOpenDynamic, adLockOptimistic
   MsgBox rs!cveempresa, vbOKOnly, ""
   While Not rs.EOF
         var_descuento_1 = 0
         var_descuento_2 = 0
         var_descuento_3 = 0
         var_porcentaje_iva = 0
         var_tipo_cambio = 1
         var_descuento_1 = IIf(IsNull(rs!descuento), 0, rs!descuento)
         var_descuento_2 = IIf(IsNull(rs!dctoptopag), 0, rs!dctoptopag)
         var_descuento_3 = 0
         var_porcentaje_iva = rs!iva
         var_tipo_cambio = IIf(IsNull(rs!comision), 1, rs!comision)
        
         var_numero = IIf(IsNull(rs!numfactura), "0", rs!numfactura)
         If IsNumeric(var_numero) Then
            var_serie = "X"
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
         End If
         Cadena = "update TB_TEM_MIGRACION_CARTERA set FLOA_TEM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_TEM_dESCUENTO_2 =  " + CStr(var_descuento_2) + ", FLOA_TEM_PORCENTAJE_IVA = " + CStr(var_porcentaje_iva) + ", FLOA_TEM_TIPO_CAMBIO = " + CStr(var_tipo_cambio) + " WHERE VCHA_TEM_CLAVE_DOCUMENTO = 'FA' AND VCHA_TEM_SERIE_DOCUMENTO = '" + var_serie + "' AND INTE_TEM_MUMERO_DOCUMENTO = " + CStr(var_numero_documento)
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
            var_clase_documento = SM
         Else
            var_tipo_documento = "FA"
            var_clase_documento = "FA"
         End If
         rsaux.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Cadena = Cadena + "FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE,FLOA_CAR_IMPORTE_NETO,VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO,VCHA_AUD_MAQUINA,VCHA_AUD_FECHA,FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID,FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS) values"
         Cadena = Cadena + "('" + rs!vcha_emp_empresa_id + "', '', '" + var_tipo_documento + "', '" + var_documento + "', '" + var_clase_documento + "', " + CStr(rs!INTE_TEM_MUMERO_DOCUMENTO) + ",'+', '', '', 0, '" + CStr(rs!DTIM_TEM_FECHA_DOCUMENTO) + "', '" + rsaux!VCHA_AGE_AGENTE_ID + "', '" + IIf(IsNull(rsaux!vcha_gac_grupo_Actual_id), "", rsaux!vcha_gac_grupo_Actual_id) + "', '" + IIf(IsNull(rsaux!vcha_gre_grupo_real_id), "", rsaux!vcha_gre_grupo_real_id) + "', '" + rsaux!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "','ESTABLECIMIENTO', " + CStr(IIf(IsNull(rsaux!inte_pla_dias), 30, rsaux!inte_pla_dias)) + "," + CStr(rs!FLOA_TEM_PORCENTAJE_IVA) + ", 0, 0, " + CStr(rs!FLOA_TEM_DESCUENTO_1) + ", " + CStr(rs!FLOA_TEM_DESCUENTO_2) + ", 0, " + CStr(var_importe_neto) + ", " + CStr(var_iva) + ", 0, 0, " + CStr(var_descuento_1) + ", " + CStr(var_descuento_2)
         Cadena = Cadena + ",0, " + CStr(var_subimporte) + ", " + CStr(rs!FLOA_TEM_IMPORTE_NETO) + ", '', 'USUARIO', 'MAQUINA', " + CStr(rs!DTIM_TEM_FECHA_CAPTURA) + ", 0,null, null, "
         Cadena = Cadena + IIf(IsNull(rsaux!vcha_mon_moneda_id), "1", rsaux!vcha_mon_moneda_id) + ", " + CStr(rs!FLOA_TEM_TIPO_CAMBIO) + ", '" + rs!VCHA_TEM_SERIE_DOCUMENTO + "', 'I')"
         rsaux2.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         rsaux3.Open "INSERT INTO TB_ESTADO_CUENTA (VCHA_EMP_EMPRESA_ID, VCHA_ECU_SERIE_CARGO, VCHA_ECU_MOVIMIENTO_CARGO, INTE_ECU_NUMERO_CARGO, FLOA_ECU_IMPORTE_CARGO, FLOA_ECU_IMPORTE_ABONO) Values ('" + rs!vcha_emp_empresa_id + "','" + rs!VCHA_TEM_SERIE_DOCUMENTO + "', '" + var_tipo_documento + "'," + CStr(rs!INTE_TEM_MUMERO_DOCUMENTO) + " , " + CStr(rs!FLOA_TEM_IMPORTE_NETO) + ", 0)", cnn, adOpenDynamic, adLockOptimistic
         rsaux.Close
         var_i = var_i + 1
         Label2 = CStr(var_i)
         Me.Refresh
         rs.MoveNext
   Wend
   rs.Close
   'hasta aqui
MsgBox "Se a terminado la migración de la cartera", vbOKOnly, "ATENCION"
End Sub

Private Sub Command1_Click()
   Cadena = "select cveempresa,cveagente,nombre,tipo,status from agentes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "insert into tb_agentes (vcha_emp_empresa_id,vcha_age_agente_id,vcha_age_nombre,VCHA_TAG_TIPOAGENTE_ID,VCHA_AGE_ESTATUS) values ('" + Trim(rsaux!cveempresa) + "','" + Trim(rsaux!cveagente) + "','" + Trim(rsaux!nombre) + "','" + rsaux!tipo + "','" + rsaux!status + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command10_Click()
   Cadena = "select cvetitular,nombre,cvepais as pais,cveestado as estado,direccion,telefono from titular"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 1
   clave_string = ""
   While Not rsaux.EOF
      rsaux2.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + rsaux!cvetitular + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux2.EOF Then
         rs.Open "SELECT MAX(VCHA_TIT_TITULAR_ID) FROM TB_TITULARES", cnn, adOpenDynamic, adLockOptimistic
         clave_string = IIf(IsNull(rs(0).Value), "0", Mid(Trim(rs(0).Value), 2, 10))
         rs.Close
         clave_numero = CInt(clave_string)
         clave_string = CStr(clave_numero + 1)
         If Len(Trim(clave_string)) = 1 Then
              clave_string = "T000000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 2 Then
            clave_string = "T00000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 3 Then
            clave_string = "T0000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 4 Then
            clave_string = "T000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 5 Then
            clave_string = "T00000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 6 Then
            clave_string = "T0000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 7 Then
            clave_string = "T000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 8 Then
            clave_string = "T00" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 9 Then
            clave_string = "T0" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 10 Then
            clave_string = "T" + Trim(clave_string)
         End If
         rs.Open "insert into tb_titulares (vcha_tit_titular_id,vcha_tit_nombre,VCHA_pai_pais_ID,VCHA_est_estado_id,vcha_tit_domicilio,vcha_tit_telefono, VCHA_TIT_TITULAR_ANTERIOR_ID) values ('" + clave_string + "','" + UCase(Trim(rsaux!nombre)) + "','" + Trim(rsaux!PAIS) + "','" + rsaux!ESTADO + "','" + UCase(Trim(rsaux!DIRECCION)) + "','" + Trim(rsaux!telefono) + "', '" + Trim(rsaux!cvetitular) + "')", cnn, adOpenDynamic, adLockOptimistic
      End If
      rsaux.MoveNext
      rsaux2.Close
   Wend
   rsaux.Close

   Cadena = "select cvecliente, cvetitular, cvefamcte, cvefamdcto from claves"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      var_clave_titular = ""
      var_clave_grupo_real = ""
      var_clave_grupo_actual = ""
      rsaux2.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + rsaux!cvetitular + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_clave_titular = rsaux2!vcha_tit_titular_id
      End If
      rsaux3.Open "select * from tb_gruposreales where vcha_gre_grupo_real_anterior_id = '" + Trim(rsaux!cvefamcte) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux3.EOF Then
         var_clave_grupo_real = rsaux3!vcha_gre_grupo_real_id
      End If
      rsaux1.Open "select * from tb_gruposactuales where vcha_gac_grupo_Actual_anterior_id = '" + Trim(rsaux!cvefamdcto) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_grupo_actual = rsaux1!vcha_gac_grupo_Actual_id
      End If
      If Trim(var_clave_titular) <> "" Then
         rs.Open "update tb_titulares set vcha_gre_grupo_real_id = '" + Trim(var_clave_grupo_real) + "' where vcha_tit_titular_id = '" + Trim(var_clave_titular) + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
      If Trim(var_clave_grupo_real) <> "" Then
         rs.Open "update tb_gruposreales set vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "', vcha_gac_grupo_actual_anterior_id = '" + Trim(rsaux!cvefamdcto) + "' where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
      rsaux3.Close
      rsaux2.Close
      rsaux1.Close
      rsaux.MoveNext
   Wend
   rsaux.Close
   

   Dim fecha As Date
   Cadena = "select cvecliente, razonsocia, fechaalta, cveagente, rfc, cvetitular, tipoclient, direccion, pais, estado, codigopost from clientes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 0
   While Not rsaux.EOF
      If Trim(rsaux!cvetitular) = "T16010" Then
         x = 1
      End If
      rsaux3.Open "select * from  tb_clientes where vcha_cli_clave_anterior_id = '" + rsaux!cvecliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux3.EOF Then
         rs.Open "SELECT MAX(vcha_cli_clave_id) FROM tb_clientes", cnn, adOpenDynamic, adLockOptimistic
         clave_string = IIf(IsNull(rs(0).Value), "0", Mid(Trim(rs(0).Value), 2, 10))
         rs.Close
         clave_numero = CInt(clave_string)
         clave_string = CStr(clave_numero + 1)
         If Len(Trim(clave_string)) = 1 Then
            clave_string = "C000000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 2 Then
            clave_string = "C00000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 3 Then
            clave_string = "C0000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 4 Then
            clave_string = "C000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 5 Then
            clave_string = "C00000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 6 Then
            clave_string = "C0000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 7 Then
            clave_string = "C000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 8 Then
            clave_string = "C00" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 9 Then
            clave_string = "C0" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 10 Then
            clave_string = "C" + Trim(clave_string)
         End If
         fecha = rsaux!fechaalta
         If fecha = "12:00:00 a.m." Then
            fecha = Date
         End If
         var_clave_agente = ""
         var_clave_titular = ""
         rsaux1.Open "select * from tb_agentes where vcha_age_agente_anterior_id = '" + Trim(rsaux!cveagente) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            var_clave_agente = rsaux1!VCHA_AGE_AGENTE_ID
         End If
         rsaux1.Close
         rsaux2.Open "select *  from tb_titulares where vcha_tit_titular_anterior_id = '" + Trim(rsaux!cvetitular) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            var_clave_titular = rsaux2!vcha_tit_titular_id
         End If
         rsaux2.Close
         rs.Open "insert into tb_clientes (VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE,DTIM_CLI_FECHA_CAPTURA,VCHA_AGE_AGENTE_ID,VCHA_CLI_RFC,VCHA_TIT_TITULAR_ID,CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_CLI_DIRECCION, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_CLI_CP) values ('" + clave_string + "','" + UCase(Trim(rsaux!razonsocia)) + "','" + CStr(fecha) + "','" + var_clave_agente + "','" + Trim(rsaux!rfc) + "','" + var_clave_titular + "','" + Trim(rsaux!tipoclient) + "', '" + Trim(rsaux!cvecliente) + "', '" + rsaux!DIRECCION + "', '" + rsaux!PAIS + "', '" + rsaux!ESTADO + "', '" + rsaux!CODIGOPOST + "')", cnn, adOpenDynamic, adLockOptimistic
      End If
      rsaux3.Close
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   Cadena = "select cvetitular,cvetienda,direccion,cvepais as pais,cveestado as estado,telefono from tiendas"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 0
   While Not rsaux.EOF
      rsaux2.Open "select * from tb_establecimientos where vcha_esb_establecimiento_anterior_id = '" + Trim(rsaux!cvetienda) + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux2.EOF Then
         rs.Open "SELECT MAX(vcha_Esb_establecimiento_id) FROM tb_establecimientos", cnn, adOpenDynamic, adLockOptimistic
         clave_string = IIf(IsNull(rs(0).Value), "0", Mid(Trim(rs(0).Value), 2, 10))
         rs.Close
         clave_numero = CInt(clave_string)
         clave_string = CStr(clave_numero + 1)
         clave_numero = clave_numero + 1
         clave_string = Trim(Str(clave_numero))
         If Len(Trim(clave_string)) = 1 Then
            clave_string = "E000000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 2 Then
            clave_string = "E00000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 3 Then
            clave_string = "E0000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 4 Then
            clave_string = "E000000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 5 Then
            clave_string = "E00000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 6 Then
            clave_string = "E0000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 7 Then
            clave_string = "E000" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 8 Then
            clave_string = "E00" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 9 Then
            clave_string = "E0" + Trim(clave_string)
         End If
         If Len(Trim(clave_string)) = 10 Then
            clave_string = "E" + Trim(clave_string)
         End If
         var_clave_titular = ""
         rs.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + Trim(rsaux!cvetitular) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_titular = rs!vcha_tit_titular_id
         End If
         rs.Close
         rs.Open "insert into tb_establecimientos (vcha_tit_titular_id, vcha_esb_establecimiento_id, vcha_esb_nombre, VCHA_pai_pais_ID, VCHA_est_estado_id, vcha_esb_domicilio, vcha_esb_telefono, vcha_esb_establecimiento_anterior_id) values ('" + var_clave_titular + "', '" + clave_string + "', '" + UCase(Trim(rsaux!DIRECCION)) + "', '" + Trim(rsaux!PAIS) + "','" + Trim(rsaux!ESTADO) + "','" + UCase(Trim(rsaux!DIRECCION)) + "','" + Trim(rsaux!telefono) + "', '" + Trim(rsaux!cvetienda) + "')", cnn, adOpenDynamic, adLockOptimistic
      End If
      rsaux2.Close
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   Cadena = "select cvecliente,cvetienda from detatien"
   rsaux.Open "delete from tb_detalle_establecimientos", cnn, adOpenDynamic, adLockOptimistic
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rsaux1.Open "select * from tb_clientes where vcha_cli_clave_anterior_id = '" + Trim(rsaux!cvecliente) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_cliente = rsaux1!vcha_cli_clave_id
      End If
      rsaux1.Close
      rsaux1.Open "select * from tb_establecimientos where vcha_esb_establecimiento_anterior_id = '" + Trim(rsaux!cvetienda) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_establecimiento = rsaux1!vcha_esb_establecimiento_id
      End If
      rsaux1.Close
      rs.Open "insert into tb_detalle_establecimientos (VCHA_cli_clave_id, VCHA_esb_establecimiento_id) values ('" + var_clave_cliente + "', '" + var_clave_establecimiento + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   
   Cadena = "select cveagente,tipo from agentes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "select * from tb_agentes where vcha_age_agente_anterior_id = '" + Trim(rsaux!cveagente) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_agente = rs!VCHA_AGE_AGENTE_ID
      End If
      rs.Close
      rs.Open "update tb_clientes set vcha_tcl_tipo_cliente_id = '" + Trim(rsaux!tipo) + "' where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close




   MsgBox "Proceso terminado", vbOKOnly, "ATENCION"
End Sub

Private Sub Command11_Click()
      rsaux2.Open "select * from titular", var_tabla, adOpenDynamic, adLockOptimistic
      While Not rsaux2.EOF
            rs.Open "update tb_clientes set INTE_CLI_ASIGNACION_CATALOGOS = 1 where vcha_cli_clave_anterior_id = '" + Trim(rsaux2!cvecliente) + "'", cnn, adOpenDynamic, adLockOptimistic
            rsaux2.MoveNext
      Wend
      rsaux2.Close
End Sub

Private Sub Command2_Click()
   Cadena = "select cvetitular,nombre,cvepais as pais,cveestado as estado,direccion,telefono from titular"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rsaux2.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + rsaux!cvetitular + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux2.EOF Then
         rs.Open "insert into tb_titulares (vcha_tit_titular_id,vcha_tit_nombre,VCHA_pai_pais_ID,VCHA_est_estado_id,vcha_tit_domicilio,vcha_tit_telefono) values ('" + Trim(rsaux!cvetitular) + "','" + Trim(rsaux!nombre) + "','" + Trim(rsaux!PAIS) + "','" + rsaux!ESTADO + "','" + rsaux!DIRECCION + "','" + Trim(rsaux!telefono) + "')", cnn, adOpenDynamic, adLockOptimistic
      End If
      rsaux2.Close
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command3_Click()
   Cadena = "select cvetitular,cvetienda,direccion,cvepais as pais,cveestado as estado,telefono from tiendas"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "insert into tb_establecimientos (vcha_tit_titular_id, vcha_esb_establecimiento_id, vcha_esb_nombre, VCHA_pai_pais_ID, VCHA_est_estado_id, vcha_esb_domicilio, vcha_esb_telefono) values ('" + Trim(rsaux!cvetitular) + "', '" + Trim(rsaux!cvetienda) + "', '" + Trim(rsaux!DIRECCION) + "', '" + Trim(rsaux!PAIS) + "','" + Trim(rsaux!ESTADO) + "','" + Trim(rsaux!DIRECCION) + "','" + Trim(rsaux!telefono) + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command4_Click()
   Cadena = "select cvefamcte,nombre from famictes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "insert into tb_gruposreales (VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2) values ('" + Trim(rsaux!cvefamcte) + "', '" + Trim(rsaux!nombre) + "', 0, 0)", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command5_Click()
   Cadena = "select cclafam,dnomfam from cfamcli"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "insert into tb_gruposactuales (VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GAC_NOMBRE, FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2) values ('" + Trim(rsaux!cclafam) + "', '" + Trim(rsaux!dnomfam) + "', 0, 0)", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command6_Click()
   Cadena = "select cveagente,tipo from agentes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "update tb_clientes set vcha_tcl_tipo_cliente_id = '" + Trim(rsaux!tipo) + "' where vcha_age_agente_id = '" + rsaux!cveagente + "'", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command7_Click()
' Para que funcione esto se debe de crear unta tabla que se lla ma claves en donde se saquen los campo
   'Cadena = "select cvecliente,cvetitular,cvefamcte,cvefamdcto from clientes into table claves"
   Cadena = "select cvecliente,cvetitular,cvefamcte,cvefamdcto from claves"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "update tb_titulares set vcha_gre_grupo_real_id = '" + Trim(rsaux!cvefamcte) + "' where vcha_tit_titular_id = '" + rsaux!cvetitular + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "update tb_gruposreales set vcha_gac_grupo_actual_id = '" + Trim(rsaux!cvefamdcto) + "' where vcha_gre_grupo_real_id = '" + rsaux!cvefamcte + "'", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command8_Click()
   Cadena = "select cvecliente,cvetienda from detatien"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "insert into tb_detalle_establecimientos (VCHA_cli_clave_id, VCHA_esb_establecimiento_id) values ('" + Trim(rsaux!cvecliente) + "', '" + Trim(rsaux!cvetienda) + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Command9_Click()
   Dim clave_string As String
   Dim clave_numero As Integer
   rs.Open "delete from tb_titulares", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_establecimientos", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_clientes", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_gruposactuales", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_gruposreales", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_detalle_establecimientos", cnn, adOpenDynamic, adLockOptimistic
   Cadena = "select cvefamdcto, nombre from famidcto"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 1
   clave_string = ""
   While Not rsaux.EOF
      clave_string = Trim(Str(clave_numero))
      If Len(Trim(clave_string)) = 1 Then
         clave_string = "GA00000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 2 Then
         clave_string = "GA0000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 3 Then
         clave_string = "GA000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 4 Then
         clave_string = "GA00" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 5 Then
         clave_string = "GA0" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 6 Then
         clave_string = "GA" + Trim(clave_string)
      End If
      rs.Open "insert into tb_gruposactuales (VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GAC_NOMBRE, FLOA_GAC_DESCUENTO_1, FLOA_GAC_DESCUENTO_2, FLOA_GAC_DESCUENTO_3, VCHA_GAC_GRUPO_ACTUAL_ANTERIOR_ID) values ('" + clave_string + "', '" + Trim(rsaux(1).Value) + "', 0, 0,0,'" + Trim(rsaux(0).Value) + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
      clave_numero = clave_numero + 1
   Wend
   rsaux.Close
   
   Cadena = "select cvefamcte,nombre from famictes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 1
   While Not rsaux.EOF
      clave_string = Trim(Str(clave_numero))
      If Len(Trim(clave_string)) = 1 Then
         clave_string = "GR00000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 2 Then
         clave_string = "GR0000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 3 Then
         clave_string = "GR000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 4 Then
         clave_string = "GR00" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 5 Then
         clave_string = "GR0" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 6 Then
         clave_string = "GR" + Trim(clave_string)
      End If
      rs.Open "insert into tb_gruposreales (VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2, FLOA_GRE_DESCUENTO_3, VCHA_GRE_GRUPO_REAL_ANTERIOR_ID) values ('" + clave_string + "', '" + UCase(Trim(rsaux!nombre)) + "', 0, 0, 0, '" + Trim(rsaux!cvefamcte) + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
      clave_numero = clave_numero + 1
   Wend
   rsaux.Close
   
   Cadena = "select cvetitular,nombre,cvepais as pais,cveestado as estado,direccion,telefono from titular"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 1
   clave_string = ""
   While Not rsaux.EOF
      clave_string = Trim(Str(clave_numero))
      If Len(Trim(clave_string)) = 1 Then
           clave_string = "T000000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 2 Then
         clave_string = "T00000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 3 Then
         clave_string = "T0000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 4 Then
         clave_string = "T000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 5 Then
         clave_string = "T00000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 6 Then
         clave_string = "T0000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 7 Then
         clave_string = "T000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 8 Then
         clave_string = "T00" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 9 Then
         clave_string = "T0" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 10 Then
         clave_string = "T" + Trim(clave_string)
      End If
      rs.Open "insert into tb_titulares (vcha_tit_titular_id,vcha_tit_nombre,VCHA_pai_pais_ID,VCHA_est_estado_id,vcha_tit_domicilio,vcha_tit_telefono, VCHA_TIT_TITULAR_ANTERIOR_ID) values ('" + clave_string + "','" + UCase(Trim(rsaux!nombre)) + "','" + Trim(rsaux!PAIS) + "','" + rsaux!ESTADO + "','" + UCase(Trim(rsaux!DIRECCION)) + "','" + Trim(rsaux!telefono) + "', '" + Trim(rsaux!cvetitular) + "')", cnn, adOpenDynamic, adLockOptimistic
      clave_numero = clave_numero + 1
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   Cadena = "select cvecliente, cvetitular, cvefamcte, cvefamdcto from claves"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      var_clave_titular = ""
      var_clave_grupo_real = ""
      var_clave_grupo_actual = ""
      rsaux2.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + rsaux!cvetitular + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_clave_titular = rsaux2!vcha_tit_titular_id
      End If
      rsaux3.Open "select * from tb_gruposreales where vcha_gre_grupo_real_anterior_id = '" + Trim(rsaux!cvefamcte) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux3.EOF Then
         var_clave_grupo_real = rsaux3!vcha_gre_grupo_real_id
      End If
      rsaux1.Open "select * from tb_gruposactuales where vcha_gac_grupo_Actual_anterior_id = '" + Trim(rsaux!cvefamdcto) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_grupo_actual = rsaux1!vcha_gac_grupo_Actual_id
      End If
      If Trim(var_clave_titular) <> "" Then
         rs.Open "update tb_titulares set vcha_gre_grupo_real_id = '" + Trim(var_clave_grupo_real) + "' where vcha_tit_titular_id = '" + Trim(var_clave_titular) + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
      If Trim(var_clave_grupo_real) <> "" Then
         rs.Open "update tb_gruposreales set vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "', vcha_gac_grupo_actual_anterior_id = '" + Trim(rsaux!cvefamdcto) + "' where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
      rsaux3.Close
      rsaux2.Close
      rsaux1.Close
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   Dim fecha As Date
   Cadena = "select cvecliente,razonsocia,fechaalta,cveagente,rfc,cvetitular,tipoclient,direccion, pais, estado, codigopost from clientes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 0
   While Not rsaux.EOF
      clave_numero = clave_numero + 1
      clave_string = Trim(Str(clave_numero))
      If Len(Trim(clave_string)) = 1 Then
         clave_string = "C000000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 2 Then
         clave_string = "C00000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 3 Then
         clave_string = "C0000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 4 Then
         clave_string = "C000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 5 Then
         clave_string = "C00000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 6 Then
         clave_string = "C0000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 7 Then
         clave_string = "C000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 8 Then
         clave_string = "C00" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 9 Then
         clave_string = "C0" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 10 Then
         clave_string = "C" + Trim(clave_string)
      End If
      fecha = rsaux!fechaalta
      If fecha = "12:00:00 a.m." Then
         fecha = Date
      End If
      var_clave_agente = ""
      var_clave_titular = ""
      rsaux1.Open "select * from tb_agentes where vcha_age_agente_anterior_id = '" + Trim(rsaux!cveagente) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_agente = rsaux1!VCHA_AGE_AGENTE_ID
      End If
      rsaux1.Close
      rsaux2.Open "select *  from tb_titulares where vcha_tit_titular_anterior_id = '" + Trim(rsaux!cvetitular) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_clave_titular = rsaux2!vcha_tit_titular_id
      End If
      rsaux2.Close
      rs.Open "insert into tb_clientes (VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE,DTIM_CLI_FECHA_CAPTURA,VCHA_AGE_AGENTE_ID,VCHA_CLI_RFC,VCHA_TIT_TITULAR_ID,CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_CLI_DIRECCION, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_CLI_CP) values ('" + clave_string + "','" + UCase(Trim(rsaux!razonsocia)) + "','" + CStr(fecha) + "','" + var_clave_agente + "','" + Trim(rsaux!rfc) + "','" + var_clave_titular + "','" + Trim(rsaux!tipoclient) + "', '" + Trim(rsaux!cvecliente) + "', '" + rsaux!DIRECCION + "', '" + rsaux!PAIS + "', '" + rsaux!ESTADO + "', '" + rsaux!CODIGOPOST + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   Cadena = "select cvetitular,cvetienda,direccion,cvepais as pais,cveestado as estado,telefono from tiendas"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   clave_numero = 0
   While Not rsaux.EOF
      clave_numero = clave_numero + 1
      clave_string = Trim(Str(clave_numero))
      If Len(Trim(clave_string)) = 1 Then
         clave_string = "E000000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 2 Then
         clave_string = "E00000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 3 Then
         clave_string = "E0000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 4 Then
         clave_string = "E000000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 5 Then
         clave_string = "E00000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 6 Then
         clave_string = "E0000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 7 Then
         clave_string = "E000" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 8 Then
         clave_string = "E00" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 9 Then
         clave_string = "E0" + Trim(clave_string)
      End If
      If Len(Trim(clave_string)) = 10 Then
         clave_string = "E" + Trim(clave_string)
      End If
      var_clave_titular = ""
      rs.Open "select * from tb_titulares where vcha_tit_titular_anterior_id = '" + Trim(rsaux!cvetitular) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_titular = rs!vcha_tit_titular_id
      End If
      rs.Close
      rs.Open "insert into tb_establecimientos (vcha_tit_titular_id, vcha_esb_establecimiento_id, vcha_esb_nombre, VCHA_pai_pais_ID, VCHA_est_estado_id, vcha_esb_domicilio, vcha_esb_telefono, vcha_esb_establecimiento_anterior_id) values ('" + var_clave_titular + "', '" + clave_string + "', '" + UCase(Trim(rsaux!DIRECCION)) + "', '" + Trim(rsaux!PAIS) + "','" + Trim(rsaux!ESTADO) + "','" + UCase(Trim(rsaux!DIRECCION)) + "','" + Trim(rsaux!telefono) + "', '" + Trim(rsaux!cvetienda) + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   Cadena = "select cvecliente,cvetienda from detatien"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rsaux1.Open "select * from tb_clientes where vcha_cli_clave_anterior_id = '" + Trim(rsaux!cvecliente) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_cliente = rsaux1!vcha_cli_clave_id
      End If
      rsaux1.Close
      rsaux1.Open "select * from tb_establecimientos where vcha_esb_establecimiento_anterior_id = '" + Trim(rsaux!cvetienda) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_clave_establecimiento = rsaux1!vcha_esb_establecimiento_id
      End If
      rsaux1.Close
      rs.Open "insert into tb_detalle_establecimientos (VCHA_cli_clave_id, VCHA_esb_establecimiento_id) values ('" + var_clave_cliente + "', '" + var_clave_establecimiento + "')", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   
   Cadena = "select cveagente,tipo from agentes"
   rsaux.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
      rs.Open "select * from tb_agentes where vcha_age_agente_anterior_id = '" + Trim(rsaux!cveagente) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_clave_agente = rs!VCHA_AGE_AGENTE_ID
      End If
      rs.Close
      rs.Open "update tb_clientes set vcha_tcl_tipo_cliente_id = '" + Trim(rsaux!tipo) + "' where vcha_age_agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      rsaux.MoveNext
   Wend
   rsaux.Close
   
   
   MsgBox "Se a terminado el proceso de migracion de información", vbOKOnly, "ATENCION"
End Sub

Private Sub Form_Load()
   Dim var_ruta As String
   var_ruta = "c:\sistemas\desarrollo\integral"
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_migracion)
End Sub
