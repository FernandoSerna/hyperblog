VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreportes_depositos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes depositos"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1980
      Left            =   390
      TabIndex        =   9
      Top             =   -90
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1440
         Left            =   45
         TabIndex        =   10
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   11
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "BC"
      Height          =   330
      Left            =   1155
      Picture         =   "frmreportes_depositos.frx":0000
      TabIndex        =   4
      ToolTipText     =   "Movimientos salvo buen cobro"
      Top             =   15
      Width           =   360
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "NR"
      Height          =   330
      Left            =   795
      Picture         =   "frmreportes_depositos.frx":063A
      TabIndex        =   3
      ToolTipText     =   "Depositos no relacionados"
      Top             =   15
      Width           =   360
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "RE"
      Height          =   330
      Left            =   435
      Picture         =   "frmreportes_depositos.frx":0C74
      TabIndex        =   1
      ToolTipText     =   "Depositos reasignados"
      Top             =   15
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "R"
      Height          =   330
      Left            =   75
      Picture         =   "frmreportes_depositos.frx":12AE
      TabIndex        =   0
      ToolTipText     =   "Depositos relacionados"
      Top             =   15
      Width           =   360
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6000
      Picture         =   "frmreportes_depositos.frx":18E8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   345
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   45
      TabIndex        =   8
      Top             =   375
      Width           =   6300
   End
   Begin VB.Frame Frame1 
      Caption         =   " Agente "
      Height          =   1380
      Left            =   105
      TabIndex        =   2
      Top             =   435
      Width           =   6225
      Begin VB.TextBox txt_nombre_agente 
         Height          =   360
         Left            =   1305
         TabIndex        =   7
         Top             =   570
         Width           =   4815
      End
      Begin VB.TextBox txt_agente 
         Height          =   360
         Left            =   150
         TabIndex        =   6
         Top             =   570
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmreportes_depositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If Me.txt_agente <> "" Then
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cnn.BeginTrans
         rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_DEPOSITOS_RELACIONADOS", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rsaux.Close
         rsaux.Open "INSERT INTO TB_TEMP_REPORTE_DEPOSITOS_RELACIONADOS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         If var_empresa = "18" Then
            var_clave = "03" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
         Else
            If var_empresa = "16" Then
               var_clave = "04" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
            Else
               var_clave = "02" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3) '
            End If
         End If

         
          'var_cadena = " select * from (select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO,vcha_car_num_docum, TO_CHAR(RCO.VCHA_CAR_DOCUMENTO) as VCHA_CAR_DOCUMENTO , TO_CHAR(RCO.VCHA_RCO_FOLIO) as vcha_Rco_folio, TO_CHAR(DEP.CUENTA) as cuenta, DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO, TO_CHAR(DEP.NO_AUTORIZACION) AS NO_AUTORIZACION,VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo From tb_cargo, vw_depositos_banc DEP, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE, TIT.VCHA_TIT_TITULAR_ID FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT  WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID) TIT, VW_TB_RC@msqlsiddist rco Where DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND TO_CHAR(rco.inte_car_numero) = TO_CHAR(vcha_car_num_docum)"
          'var_cadena = var_cadena + " AND TO_CHAR(rco.inte_rco_numero_deposito) = TO_CHAR(DEP.FOLIO) AND TO_CHAR(RCO.VCHA_TIT_TITULAR_ID) = TO_CHAR(TIT.VCHA_TIT_TITULAR_ID) AND TO_CHAR(TIT.VCHA_CLI_REFERENCIA) = TO_CHAR(DEP.referencia) AND TO_CHAR(DEP.NO_AUTORIZACION(+)) = TO_CHAR(inte_car_abono_id) AND VCHA_CAR_TIPO_DOCUMENTO = 'RC' Union All select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO,vcha_car_num_docum ,'REASIGNACION','---', TO_CHAR(DEP.CUENTA), DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO, TO_CHAR(DEP.NO_AUTORIZACION), VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo from tb_cargo, vw_depositos_banc DEP, "
          'var_cadena = var_cadena + " (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE, TIT.VCHA_TIT_TITULAR_ID FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND vcha_cli_referencia = '020160003107') AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT Where DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND VCHA_CAR_TIPO_DOCUMENTO = 'TA' AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia and TO_NUMBER(DEP.NO_AUTORIZACION(+)) = TO_NUMBER(inte_car_abono_id) Union All"
          'var_cadena = var_cadena + " select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO ,vcha_car_num_docum ,'ANULACION BANCARIA','---', DEP.CUENTA, DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO,  TO_CHAR(DEP.NO_AUTORIZACION), VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo from tb_cargo, vw_depositos_banc DEP, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM  tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(vcha_cli_referencia) = 12) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT Where"
          'var_cadena = var_cadena + " DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND VCHA_CAR_TIPO_DOCUMENTO IN ('ABM','ABA') AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia and DEP.NO_AUTORIZACION(+) = inte_car_abono_id Union All select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO ,VCHA_CAR_TIPO_DOCUMENTO ,'CONTRALORIA','---', DEP.CUENTA, DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO,  TO_CHAR(DEP.NO_AUTORIZACION), VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo from tb_cargo, vw_depositos_banc DEP, "
          'var_cadena = var_cadena + " (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM  tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(vcha_cli_referencia) = 12) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT Where DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND VCHA_CAR_TIPO_DOCUMENTO NOT IN ('RC','TA','ABM','ABA') AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia and DEP.NO_AUTORIZACION(+) = inte_car_abono_id ) WS order by ws.NO_AUTORIZACION, ws.fecha_autorizacion "
                                     
                                               
          var_cadena = "select * from (select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO,vcha_car_num_docum, TO_CHAR(RCO.VCHA_CAR_DOCUMENTO) as VCHA_CAR_DOCUMENTO "
          var_cadena = var_cadena + " , TO_CHAR(RCO.VCHA_RCO_FOLIO) as VCHA_RCO_FOLIO, TO_CHAR(DEP.CUENTA) as CUENTA, DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO, TO_CHAR(DEP.NO_AUTORIZACION) AS NO_AUTORIZACION,VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo, TB_CARGO.date_car_fecha_cargo AS FECHA_APLICACION From tb_cargo, vw_depositos_banc DEP, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE, TIT.VCHA_TIT_TITULAR_ID FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT  WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID) TIT, VW_TB_RC@msqlsiddist rco Where DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND TO_CHAR(rco.inte_car_numero) = TO_CHAR(vcha_car_num_docum)"
          var_cadena = var_cadena + " AND TO_CHAR(rco.inte_rco_numero_deposito) = TO_CHAR(DEP.FOLIO) AND TO_CHAR(RCO.VCHA_TIT_TITULAR_ID) = TO_CHAR(TIT.VCHA_TIT_TITULAR_ID) AND TO_CHAR(TIT.VCHA_CLI_REFERENCIA) = TO_CHAR(DEP.referencia) AND TO_CHAR(DEP.NO_AUTORIZACION(+)) = TO_CHAR(inte_car_abono_id) AND VCHA_CAR_TIPO_DOCUMENTO = 'RC' Union All select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO,vcha_car_num_docum ,'REASIGNACION','---', TO_CHAR(DEP.CUENTA), DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO, TO_CHAR(DEP.NO_AUTORIZACION), VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo, TB_CARGO.date_car_fecha_cargo AS FECHA_APLICACION from tb_cargo, vw_depositos_banc DEP, "
          var_cadena = var_cadena + " (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE, TIT.VCHA_TIT_TITULAR_ID FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND vcha_cli_referencia = '020160003107') AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT Where DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND VCHA_CAR_TIPO_DOCUMENTO = 'TA' AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia and TO_NUMBER(DEP.NO_AUTORIZACION(+)) = TO_NUMBER(inte_car_abono_id) Union All"
          var_cadena = var_cadena + " select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO ,vcha_car_num_docum ,'ANULACION BANCARIA','---', DEP.CUENTA, DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO,  TO_CHAR(DEP.NO_AUTORIZACION), VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo, TB_CARGO.date_car_fecha_cargo AS FECHA_APLICACION from tb_cargo, vw_depositos_banc DEP, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM  tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(vcha_cli_referencia) = 12) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT Where"
          var_cadena = var_cadena + " DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND VCHA_CAR_TIPO_DOCUMENTO IN ('ABM','ABA') AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia and DEP.NO_AUTORIZACION(+) = inte_car_abono_id Union All select TIT.VCHA_TIT_NOMBRE, DEP.REFERENCIA, DEP.FECHA_DEPOSITO, DEP.FECHA_AUTORIZACION, DEP.IMPORTE, NUMB_CAR_IMPORTE AS IMPORTE_CARGO ,VCHA_CAR_TIPO_DOCUMENTO ,'CONTRALORIA','---', DEP.CUENTA, DEP.DIVISA, DEP.ORIGEN, DEP.FOLIO,  TO_CHAR(DEP.NO_AUTORIZACION), VCHA_CAR_TIPO_DOCUMENTO,DEP.DESCRIPCION, dep.recibo, TB_CARGO.date_car_fecha_cargo AS FECHA_APLICACION from tb_cargo, vw_depositos_banc DEP,"
          var_cadena = var_cadena + " (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM  tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(vcha_cli_referencia) = 12) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT Where DEP.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND LENGTH(VCHA_CLI_REFERENCIA) = 12) AND VCHA_CAR_TIPO_DOCUMENTO NOT IN ('RC','TA','ABM','ABA') AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia and DEP.NO_AUTORIZACION(+) = inte_car_abono_id ) WS order by ws.NO_AUTORIZACION, ws.fecha_autorizacion"
         
         
         'MsgBox var_cadena
         Text1 = var_cadena
         rsaux.Open var_cadena, cnnoracle_2, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_cadena = "INSERT INTO TB_tEMP_REPORTE_DEPOSITOS_RELACIONADOS (INTE_TEM_CONSECUTIVO, VCHA_TIT_NOMBRE, REFERENCIA, FECHA_DEPOSITO, FECHA_AUTORIZACION, IMPORTE, IMPORTE_CARGO, VCHA_CAR_NUM_DOCUM, VCHA_CAR_DOCUM, VCHA_RCO_FOLIO, CUENTA, DIVISA, ORIGEN, FOLIO, NO_AUTORIZACION, VCHA_CAR_TIPO_DOCUMENTO, DESCRIPCION, RECIBO, FECHA_APLICACION) "
               
               var_dia = CStr(Day(CDate(rsaux!FECHA_DEPOSITO)))
               var_mes = CStr(Month(CDate(rsaux!FECHA_DEPOSITO)))
               var_año = CStr(Year(CDate(rsaux!FECHA_DEPOSITO)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_deposito = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               var_dia = CStr(Day(CDate(rsaux!FECHA_AUTORIZACION)))
               var_mes = CStr(Month(CDate(rsaux!FECHA_AUTORIZACION)))
               var_año = CStr(Year(CDate(rsaux!FECHA_AUTORIZACION)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_autorizacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

               
               var_dia = CStr(Day(CDate(rsaux!FECHA_APLICACION)))
               var_mes = CStr(Month(CDate(rsaux!FECHA_APLICACION)))
               var_año = CStr(Year(CDate(rsaux!FECHA_APLICACION)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_aplicacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

               
               var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rsaux!VCHA_TIT_NOMBRE), "", rsaux!VCHA_TIT_NOMBRE) + "', '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "', " + var_fecha_deposito + ", " + var_fecha_autorizacion + ", " + CStr(IIf(IsNull(rsaux!Importe), 0, rsaux!Importe)) + ", " + CStr(IIf(IsNull(rsaux!importe_Cargo), 0, rsaux!importe_Cargo)) + ", '" + IIf(IsNull(rsaux!vcha_Car_num_docum), "", rsaux!vcha_Car_num_docum) + "', '" + IIf(IsNull(rsaux!vcha_car_documento), "", rsaux!vcha_car_documento) + "', '" + IIf(IsNull(rsaux!vcha_Rco_folio), "", rsaux!vcha_Rco_folio) + "', '" + IIf(IsNull(rsaux!CUENTA), "", rsaux!CUENTA) + "', '" + IIf(IsNull(rsaux!DIVISA), "", rsaux!DIVISA) + "', '" + IIf(IsNull(rsaux!Origen), "", rsaux!Origen) + "', '" + CStr(IIf(IsNull(rsaux!FOLIO), "", rsaux!FOLIO)) + "', '"
               var_cadena = var_cadena + CStr(IIf(IsNull(rsaux!NO_AUTORIZACION), "", rsaux!NO_AUTORIZACION)) + "', '" + IIf(IsNull(rsaux!vcha_Car_tipo_documento), "", rsaux!vcha_Car_tipo_documento) + "', '"
               var_cadena = var_cadena + IIf(IsNull(rsaux!descripcion), "", rsaux!descripcion) + "','" + IIf(IsNull(rsaux!recibo), "", rsaux!recibo) + "'," + var_fecha_aplicacion + ")"
               'var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rsaux!VCHA_TIT_NOMBRE), "", rsaux!VCHA_TIT_NOMBRE) + "', '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "', " + var_fecha_deposito + ", " + var_fecha_autorizacion + ", " + CStr(IIf(IsNull(rsaux!Importe), 0, rsaux!Importe)) + ", '" + IIf(IsNull(rsaux!CUENTA), "", rsaux!CUENTA) + "', '" + IIf(IsNull(rsaux!DIVISA), "", rsaux!DIVISA) + "', '" + IIf(IsNull(rsaux!Origen), "", rsaux!Origen) + "', '" + IIf(IsNull(rsaux!FOLIO), "", rsaux!FOLIO) + "', " + CStr(IIf(IsNull(rsaux!NO_AUTORIZACION), 0, rsaux!NO_AUTORIZACION)) + ", '" + IIf(IsNull(rsaux!VCHA_CAR_NUM_DOCUM), "", rsaux!VCHA_CAR_NUM_DOCUM) + "')"
               'MsgBox var_cadena
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rsaux.MoveNext
         Wend
         rsaux.Close
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_DEPOSITOS_RELACIONADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_tit_nombre is null", cnn, adOpenDynamic, adLockOptimistic
         
       
         'Set reporte = appl.OpenReport(App.Path + "\rep_depositos_relacionados.rpt")
         'reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEPOSITOS_RELACIONADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         'frmvistasprevias.cr.ReportSource = reporte
         'For ntablas = 1 To reporte.Database.Tables.Count
         '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         'Next ntablas
         'frmvistasprevias.cr.ViewReport
         'frmvistasprevias.Caption = "Reporte de depositos relacionados"
         'frmvistasprevias.Show 1
         'Set reporte = Nothing
         'var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
         'If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_depositos_relacionados.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEPOSITOS_RELACIONADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_depositos_relacionados_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         'End If
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_DEPOSITOS_RELACIONADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   If Me.txt_agente <> "" Then
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cnn.BeginTrans
         rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rsaux.Close
         rsaux.Open "INSERT INTO TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         
         If var_empresa = "18" Then
            var_clave = "03" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
         Else
            If var_empresa = "16" Then
               var_clave = "04" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
            Else
               var_clave = "02" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3) '
            End If
         End If
         
         
         
         'var_cadena = " select rco.inte_rco_numero_deposito, DEP.VCHA_MOV_FOLIO_AUTORIZACION, TIT.VCHA_TIT_NOMBRE, DEP.VCHA_MOV_REFERENCIA_REAL, DEP.date_mov_fecha_operacion, DEP.DATE_ABO_FECHA_DEPOSITO "
         'var_cadena = var_cadena + " ,DEP.DATE_ABO_FECHA_AUTORIZACION, DEP.VCHA_MOV_IMPORTE_ABONO, NUMB_CAR_IMPORTE AS IMPORTE_CARGO,vcha_car_num_docum, RCO.VCHA_CAR_DOCUMENTO, RCO.VCHA_RCO_FOLIO, DEP.VCHA_CUE_CUENTA_ID, DEP.VCHA_MOV_MONEDA, DEP.VCHA_ORIGEN, DEP.VCHA_MOV_FOLIO_AUTORIZACION, DEP.VCHA_ABO_ABONO_ID,VCHA_ABO_TIPO_DOCUMENTO from tb_cargo, VW_DEPOSITOS_REASIGNADOS dep, VW_TB_RC@msqlsiddist rco, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT"
         'var_cadena = var_cadena + " where dep.VCHA_ABO_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null) AND TIT.VCHA_CLI_REFERENCIA = DEP.VCHA_ABO_REFERENCIA and rco.inte_car_numero = vcha_car_num_docum and rco.inte_rco_numero_deposito = DEP.VCHA_ABO_CLAVE_DOCUMENTO and DEP.vcha_abo_abono_id(+) = inte_car_abono_id order by dep.vcha_abo_abono_id, dep.DATE_ABO_FECHA_DEPOSITO"
         
         var_cadena = "select rco.inte_rco_numero_deposito, DEP.VCHA_MOV_FOLIO_AUTORIZACION, TIT.VCHA_TIT_NOMBRE, DEP.VCHA_MOV_REFERENCIA_REAL, DEP.date_mov_fecha_operacion, "
         var_cadena = var_cadena + " DEP.DATE_ABO_FECHA_AUTORIZACION, DEP.VCHA_MOV_IMPORTE_ABONO, NUMB_CAR_IMPORTE AS IMPORTE_CARGO,vcha_car_num_docum, RCO.VCHA_CAR_DOCUMENTO, RCO.VCHA_RCO_FOLIO, DEP.VCHA_CUE_CUENTA_ID, DEP.VCHA_MOV_MONEDA, DEP.VCHA_ORIGEN, DEP.VCHA_MOV_FOLIO_AUTORIZACION, DEP.VCHA_ABO_ABONO_ID,VCHA_ABO_TIPO_DOCUMENTO from tb_cargo, VW_DEPOSITOS_REASIGNADOS dep, VW_TB_RC@msqlsiddist rco, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT"
         var_cadena = var_cadena + " where dep.VCHA_ABO_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null) AND TIT.VCHA_CLI_REFERENCIA = DEP.VCHA_ABO_REFERENCIA and rco.inte_car_numero = vcha_car_num_docum and rco.inte_rco_numero_deposito = DEP.VCHA_ABO_CLAVE_DOCUMENTO and DEP.vcha_abo_abono_id(+) = inte_car_abono_id order by dep.vcha_abo_abono_id, dep.DATE_ABO_FECHA_DEPOSITO"
         
         MsgBox var_cadena
         rsaux.Open var_cadena, cnnoracle_2, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_cadena = "INSERT INTO TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS (INTE_TEM_CONSECUTIVO, INTE_RCO_NUMERO_DEPOSITO, VCHA_MOV_FOLIO_AUTORIZACION, VCHA_TIT_NOMBRE, VCHA_ABO_REFERENCIA, DATE_ABO_FECHA_DEPOSITO, DATE_ABO_FECHA_AUTORIZACION, NUMB_ABO_IMPORTE, IMPORTE_CARGO, VCHA_CAR_NUM_DOCUM, VCHA_CAR_DOCUMENTO, VCHA_RCO_FOLIO, VCHA_CUE_CUENTA_ID, VCHA_MOV_MONEDA, VCHA_ORIGEN, VCHA_MOV_FOLIO_AUTORIZACION_1, VCHA_ABO_ABONO_ID, VCHA_ABO_TIPO_DOCUMENTO) "
               
               var_dia = CStr(Day(CDate(rsaux!date_mov_fecha_operacion)))
               var_mes = CStr(Month(CDate(rsaux!date_mov_fecha_operacion)))
               var_año = CStr(Year(CDate(rsaux!date_mov_fecha_operacion)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_deposito = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               var_dia = CStr(Day(CDate(rsaux!date_abo_fecha_autorizacion)))
               var_mes = CStr(Month(CDate(rsaux!date_abo_fecha_autorizacion)))
               var_año = CStr(Year(CDate(rsaux!date_abo_fecha_autorizacion)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_autorizacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

               
               
               var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", " + CStr(IIf(IsNull(rsaux!inte_rco_numero_deposito), 0, rsaux!inte_rco_numero_deposito)) + ",'" + CStr(IIf(IsNull(rsaux!VCHA_MOV_FOLIO_AUTORIZACION), "", rsaux!VCHA_MOV_FOLIO_AUTORIZACION)) + "','" + IIf(IsNull(rsaux!VCHA_TIT_NOMBRE), "", rsaux!VCHA_TIT_NOMBRE) + "', '" + IIf(IsNull(rsaux!VCHA_MOV_REFERENCIA_REAL), "", rsaux!VCHA_MOV_REFERENCIA_REAL) + "',  " + var_fecha_deposito + "," + var_fecha_autorizacion + "," + CStr(IIf(IsNull(rsaux!VCHA_MOV_IMPORTE_ABONO), 0, rsaux!VCHA_MOV_IMPORTE_ABONO)) + ", " + CStr(IIf(IsNull(rsaux!importe_Cargo), 0, rsaux!importe_Cargo)) + ",'" + IIf(IsNull(rsaux!vcha_Car_num_docum), "", rsaux!vcha_Car_num_docum) + "',"
               var_cadena = var_cadena + "'" + IIf(IsNull(rsaux!vcha_car_documento), "", rsaux!vcha_car_documento) + "','" + IIf(IsNull(rsaux!vcha_Rco_folio), "", rsaux!vcha_Rco_folio) + "','" + IIf(IsNull(rsaux!VCHA_CUE_CUENTA_ID), "", rsaux!VCHA_CUE_CUENTA_ID) + "', '" + IIf(IsNull(rsaux!VCHA_MOV_MONEDA), "", rsaux!VCHA_MOV_MONEDA) + "','" + IIf(IsNull(rsaux!VCHA_ORIGEN), "", rsaux!VCHA_ORIGEN) + "', '" + IIf(IsNull(rsaux!VCHA_MOV_FOLIO_AUTORIZACION), "", rsaux!VCHA_MOV_FOLIO_AUTORIZACION) + "','" + CStr(IIf(IsNull(rsaux!vcha_abo_abono_id), "", rsaux!vcha_abo_abono_id)) + "', '" + IIf(IsNull(rsaux!vcha_abo_tipo_documento), "", rsaux!vcha_abo_tipo_documento) + "')"
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rsaux.MoveNext
         Wend
         rsaux.Close
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_tit_nombre is null", cnn, adOpenDynamic, adLockOptimistic
         
       
         'Set reporte = appl.OpenReport(App.Path + "\rep_depositos_reasignados.rpt")
         'reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         'frmvistasprevias.cr.ReportSource = reporte
         'For ntablas = 1 To reporte.Database.Tables.Count
         '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         'Next ntablas
         'frmvistasprevias.cr.ViewReport
         'frmvistasprevias.Caption = "Reporte de depositos reasignados"
         'frmvistasprevias.Show 1
         'Set reporte = Nothing
         'var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
         'If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_depositos_reasignados.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_depositos_reasignados_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         'End If
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_DEPOSITOS_REASIGNADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command3_Click()
   If Me.txt_agente <> "" Then
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cnn.BeginTrans
         rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rsaux.Close
         rsaux.Open "INSERT INTO TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
                  If var_empresa = "18" Then
            var_clave = "03" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
         Else
            If var_empresa = "16" Then
               var_clave = "04" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
            Else
               var_clave = "02" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3) '
            End If
         End If
         
         
         
         'var_cadena = " SELECT TIT.VCHA_TIT_NOMBRE, VCHA_ABO_REFERENCIA, NUMB_ABO_IMPORTE AS MPORTE_ORIGINAL, APLICADO,SALDO, DATE_ABO_FECHA_DEPOSITO AS FECHA_DE_DEPOSITO, DATE_ABO_FECHA_AUTORIZACION AS FECHA_DE_AUTORIZACION From (select AGE.VCHA_AGE_NOMBRE, z.VCHA_ABO_REFERENCIA, z.NUMB_ABO_IMPORTE, z.DATE_ABO_FECHA_DEPOSITO, z.DATE_ABO_FECHA_AUTORIZACION, 0 as APLICADO, z.NUMB_ABO_IMPORTE as SALDO From (select * From tb_abono Where inte_abo_usedbyoldrc is null and vcha_abo_tipo_documento = 'DR' and vcha_abo_abono_id not in (select inte_car_abono_id from tb_cargo where inte_car_abono_id is not null) and vcha_abo_abono_id not in (select a.vcha_abo_abono_id from tb_abono a, tb_fichas_depo b, tb_cargo c, tb_movimiento d where b.vcha_fichdepo_estatus = 'REASIGNADO' and a.vcha_abo_tipo_documento = 'DR' and d.vcha_mov_clave = b.vcha_fichdepo_mov_banco and c.vcha_car_num_docum = b.vcha_fich_depo_ficha_id and c.vcha_car_tipo_documento = 'TA' and a.vcha_abo_clave_documento = d.vcha_mov_folio_autorizacion) ) z,"
         'var_cadena = var_cadena + " (select distinct(cli.vcha_cli_referencia), cli.vcha_age_agente_id from tb_clientes@msqlsiddist cli where cli.vcha_cli_referencia is not null OR LENGTH(CLI.VCHA_CLI_REFERENCIA) != 12) cli, tb_agentes@msqlsiddist age Where cli.VCHA_CLI_REFERENCIA = z.VCHA_ABO_REFERENCIA and cli.vcha_age_agente_id = age.vcha_age_agente_id AND AGE.vcha_age_agente_id = '" + Me.txt_agente + "' Union all select age.vcha_age_nombre, z.*, (z.numb_abo_importe-z.aplicado) as saldo From (select distinct(cli.vcha_cli_referencia), cli.vcha_age_agente_id from tb_clientes@msqlsiddist cli where cli.vcha_cli_referencia is not null or length(cli.vcha_cli_referencia)!=12) cli, tb_agentes@msqlsiddist age, (select a.vcha_abo_referencia, a.numb_abo_importe, a.date_abo_fecha_deposito, a.date_abo_fecha_autorizacion, sum(b.numb_car_importe) as aplicado from tb_abono a, tb_cargo b Where a.vcha_abo_abono_id = b.inte_car_abono_id"
         'var_cadena = var_cadena + " group by a.vcha_abo_referencia, a.numb_abo_importe, a.date_abo_fecha_deposito, a.date_abo_fecha_autorizacion) z Where cli.VCHA_CLI_REFERENCIA = z.VCHA_ABO_REFERENCIA and cli.vcha_age_agente_id = age.vcha_age_agente_id and (z.numb_abo_importe-z.aplicado)>0 AND AGE.vcha_age_agente_id = '" + Me.txt_agente + "') W, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230')) TIT Where"
         'var_cadena = var_cadena + " TIT.VCHA_CLI_REFERENCIA = W.VCHA_ABO_REFERENCIA"
         
         
         
          var_cadena = " SELECT TIT.VCHA_TIT_NOMBRE, TIT.VCHA_CLI_REFERENCIA, W.NUMB_ABO_IMPORTE AS MPORTE_ORIGINAL, W.APLICADO,SALDO, W.DATE_ABO_FECHA_DEPOSITO AS FECHA_DE_DEPOSITO, W.date_abo_fecha_autorizacion AS FECHA_DE_AUTORIZACION From (select AGE.VCHA_AGE_NOMBRE, z.VCHA_ABO_REFERENCIA, z.NUMB_ABO_IMPORTE, z.DATE_ABO_FECHA_DEPOSITO, TO_CHAR(z.date_abo_fecha_autorizacion, 'DD-MON-YYYY HH24:MI:SS') as date_abo_fecha_autorizacion, 0 as APLICADO, z.NUMB_ABO_IMPORTE as SALDO From (select * From tb_abono Where inte_abo_usedbyoldrc is null and vcha_abo_abono_id not in (select inte_car_abono_id from tb_cargo where inte_car_abono_id is not null) "
          var_cadena = var_cadena + " and vcha_abo_abono_id not in (select a.vcha_abo_abono_id from tb_abono a, tb_fichas_depo b, tb_cargo c, tb_movimiento d where b.vcha_fichdepo_estatus = 'REASIGNADO' and a.vcha_abo_tipo_documento = 'DR'and d.vcha_mov_clave = b.vcha_fichdepo_mov_banco and c.vcha_car_num_docum = b.vcha_fich_depo_ficha_id and c.vcha_car_tipo_documento = 'TA' and a.vcha_abo_clave_documento = d.vcha_mov_folio_autorizacion and a.vcha_abo_abono_id = c.inte_car_abono_id) ) z, (select distinct(cli.vcha_cli_referencia), cli.vcha_age_agente_id from tb_clientes@msqlsiddist cli where cli.vcha_cli_referencia is not null and LENGTH(CLI.VCHA_CLI_REFERENCIA) = 12) cli, tb_agentes@msqlsiddist age Where CLI.VCHA_CLI_REFERENCIA = z.VCHA_ABO_REFERENCIA and cli.vcha_age_agente_id = age.vcha_age_agente_id AND AGE.vcha_age_agente_id = '" + Me.txt_agente + "' Union All select age.vcha_age_nombre, z.*, (z.numb_abo_importe-z.aplicado) as saldo From"
          var_cadena = var_cadena + " (select distinct(cli.vcha_cli_referencia), cli.vcha_age_agente_id from tb_clientes@msqlsiddist cli where cli.vcha_cli_referencia is not null AND length(cli.vcha_cli_referencia) = 12) cli, tb_agentes@msqlsiddist age, (select a.vcha_abo_referencia, a.numb_abo_importe, a.date_abo_fecha_deposito, TO_CHAR(a.date_abo_fecha_autorizacion, 'DD-MON-YYYY HH24:MI:SS')as date_abo_fecha_autorizacion , sum(b.numb_car_importe) as aplicado from tb_abono a, tb_cargo b Where a.vcha_abo_abono_id = b.inte_car_abono_id group by a.vcha_abo_referencia, a.numb_abo_importe, a.date_abo_fecha_deposito, a.date_abo_fecha_autorizacion) z Where CLI.VCHA_CLI_REFERENCIA = z.VCHA_ABO_REFERENCIA and cli.vcha_age_agente_id = age.vcha_age_agente_id and (z.numb_abo_importe-z.aplicado) > 1 AND AGE.vcha_age_agente_id = '" + Me.txt_agente + "') W, ( SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT"
          var_cadena = var_cadena + " Where TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null AND length(cli.vcha_cli_referencia) = 12) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230')) TIT Where TIT.VCHA_CLI_REFERENCIA = W.VCHA_ABO_REFERENCIA"
          rsaux.Open var_cadena, cnnoracle_2, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_cadena = "INSERT INTO TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS (INTE_TEM_CONSECUTIVO, VCHA_TIT_NOMBRE, VCHA_ABO_REFERENCIA, NUMB_ABO_IMPORTE, APLICADO, SALDO, DATE_ABO_FECHA_DEPOSITO, DATE_ABO_FECHA_AUTORIZACION) "
               
               var_fecha_deposito = CStr(rsaux!FECHA_de_DEPOSITO)
               'var_dia = CStr(Day(CDate(rsaux!FECHA_de_DEPOSITO)))
               'var_mes = CStr(Month(CDate(rsaux!FECHA_de_DEPOSITO)))
               'var_año = CStr(Year(CDate(rsaux!FECHA_de_DEPOSITO)))
               'If Len(Trim(var_dia)) = 1 Then
               '   var_dia = "0" + var_dia
               'End If
               'If Len(Trim(var_mes)) = 1 Then
               '   var_mes = "0" + var_mes
               'End If
               'var_fecha_deposito = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               'var_dia = CStr(Day(CDate(rsaux!FECHA_de_AUTORIZACION)))
               'var_mes = CStr(Month(CDate(rsaux!FECHA_de_AUTORIZACION)))
               'var_año = CStr(Year(CDate(rsaux!FECHA_de_AUTORIZACION)))
               'If Len(Trim(var_dia)) = 1 Then
               '   var_dia = "0" + var_dia
               'End If
               'If Len(Trim(var_mes)) = 1 Then
               '   var_mes = "0" + var_mes
               'End If
               'var_fecha_autorizacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

               
               var_fecha_autorizacion = CStr(rsaux!FECHA_de_AUTORIZACION)
               
               var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rsaux!VCHA_TIT_NOMBRE), "", rsaux!VCHA_TIT_NOMBRE) + "', '" + IIf(IsNull(rsaux!VCHA_cli_REFERENCIA), "", rsaux!VCHA_cli_REFERENCIA) + "' , " + CStr(IIf(IsNull(rsaux!mporte_original), "", rsaux!mporte_original)) + ",   " + CStr(IIf(IsNull(rsaux!APLICADO), 0, rsaux!APLICADO)) + ", " + CStr(IIf(IsNull(rsaux!SALDO), 0, rsaux!SALDO)) + ",'" + var_fecha_deposito + "', '" + var_fecha_autorizacion + "')"
               'MsgBox var_cadena
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rsaux.MoveNext
         Wend
         rsaux.Close
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_tit_nombre is null", cnn, adOpenDynamic, adLockOptimistic
         
       
         Set reporte = appl.OpenReport(App.Path + "\rep_depositos_no_Relacionados.rpt")
         reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de depositos no relacionados"
         frmvistasprevias.Show 1
         Set reporte = Nothing
             var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_depositos_no_Relacionados.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_depositos_no_relacionados_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_DEPOSITOS_NO_RELACIONADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command4_Click()
   If Me.txt_agente <> "" Then
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cnn.BeginTrans
         rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rsaux.Close
         rsaux.Open "INSERT INTO TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         
         If var_empresa = "18" Then
            var_clave = "03" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
         Else
            If var_empresa = "16" Then
               var_clave = "04" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3)
            Else
               var_clave = "02" + Mid(rs!VCHA_AGE_AGENTE_ID, 3, 3) '
            End If
         End If
         
         'var_cadena = "select tit.VCHA_TIT_NOMBRE, dep.""TIENDA/AGENTE"" as tienda_agente, dep.ABONO, dep.MONEDA, dep.REFERENCIA,dep.ORIGEN, dep.""FECHA DE OPERACION"" as fecha_de_operacion, dep.""FECHA DE REGISTRO"" as fecha_De_registro, dep.""SE DEBE AUTORIZAR:"" as se_debe_autorizar from VW_AUTORIZAR_CHQ_DET dep, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM TB_CLIENTES@DISTRIBUCION CLI, TB_TITULARES@DISTRIBUCION TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA LIKE '" + var_clave + "%' AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT where dep.referencia like '" + var_clave + "%'  AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia"
         var_cadena = "select tit.VCHA_TIT_NOMBRE, dep.""TIENDA/AGENTE"" as tienda_agente, dep.ABONO, dep.MONEDA, dep.REFERENCIA, dep.Origen, dep.""FECHA DE OPERACION"" as fecha_de_operacion, dep.""FECHA DE REGISTRO"" as fecha_De_regsitro, dep.""SE DEBE AUTORIZAR:"" as se_debe_autorizar from VW_AUTORIZAR_CHQ_DET dep, (SELECT DISTINCT(CLI.VCHA_CLI_REFERENCIA), TIT.VCHA_TIT_NOMBRE FROM tb_clientes@msqlsiddist CLI, TB_TITULARES@msqlsiddist TIT WHERE TIT.VCHA_TIT_TITULAR_ID = CLI.VCHA_TIT_TITULAR_ID AND CLI.VCHA_CLI_REFERENCIA in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null) AND TIT.VCHA_TIT_TITULAR_ID NOT IN ('T000000230') ORDER BY CLI.VCHA_CLI_REFERENCIA) TIT where dep.referencia in (select distinct(vcha_cli_referencia) from tb_clientes@msqlsiddist where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_cli_referencia is not null)  AND TIT.VCHA_CLI_REFERENCIA = DEP.referencia"
         'MsgBox var_cadena
         rsaux.Open var_cadena, cnnoracle_2, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_cadena = "INSERT INTO TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO (INTE_TEM_CONSECUTIVO, VCHA_TIT_NOMBRE, TIENDA_AGENTE, ABONO, MONEDA, REFERENCIA, ORIGEN, FECHA_DE_OPERACION, FECHA_DE_REGISTRO, SE_DEBE_AUTORIZAR) "
               
               var_dia = CStr(Day(CDate(rsaux!fecha_de_operacion)))
               var_mes = CStr(Month(CDate(rsaux!fecha_de_operacion)))
               var_año = CStr(Year(CDate(rsaux!fecha_de_operacion)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_deposito = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               var_dia = CStr(Day(CDate(rsaux!fecha_De_regsitro)))
               var_mes = CStr(Month(CDate(rsaux!fecha_De_regsitro)))
               var_año = CStr(Year(CDate(rsaux!fecha_De_regsitro)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_autorizacion = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

               var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rsaux!VCHA_TIT_NOMBRE), "", rsaux!VCHA_TIT_NOMBRE) + "', '" + IIf(IsNull(rsaux!tienda_agente), "", rsaux!tienda_agente) + "', " + CStr(IIf(IsNull(rsaux!abono), "", rsaux!abono)) + ", '" + IIf(IsNull(moneda), "", rsaux!moneda) + "', '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "','" + IIf(IsNull(rsaux!Origen), "", rsaux!Origen) + "', " + var_fecha_deposito + ", " + var_fecha_autorizacion + ", '" + IIf(IsNull(rsaux!se_debe_autorizar), "", rsaux!se_debe_autorizar) + "')"
               rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               rsaux.MoveNext
         Wend
         rsaux.Close
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_tit_nombre is null", cnn, adOpenDynamic, adLockOptimistic
         
       
         Set reporte = appl.OpenReport(App.Path + "\rep_movimientos_salvo_buen_cobro.rpt")
         reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de movimientos salvo buen cobro"
         frmvistasprevias.Show 1
         Set reporte = Nothing
             var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_movimientos_salvo_buen_cobro.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_movimientos_salvo_buen_cobro_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         rsaux.Open "DELETE FROM TB_TEMP_REPORTE_MOVIMIENTOS_SALVO_BUEN_COBRO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Left = 2800
   Top = 2800
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_DblClick()
    Me.txt_agente = lv_lista.selectedItem
    Me.txt_nombre_agente = Me.lv_lista.selectedItem.SubItems(1)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_agente = Me.lv_lista.selectedItem
      Me.txt_nombre_agente = Me.lv_lista.selectedItem.SubItems(1)
      Me.txt_agente.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_agente.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_Empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_agente.SetFocus
   End If
End Sub

Private Sub txt_agente_LostFocus()
   If Me.txt_agente <> "" Then
     rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_aGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
     If Not rs.EOF Then
        Me.txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
     Else
        Me.txt_nombre_agente = ""
        Me.txt_agente = ""
        MsgBox "El agente no existe o no pertenece a la empresa.", vbOKOnly, "ATENCION"
     End If
     rs.Close
   Else
      Me.txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_Empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.Command1.SetFocus
   End If
End Sub
