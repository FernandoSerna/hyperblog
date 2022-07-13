VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmimportacion_clientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de clientes"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_busqueda_folio 
      Height          =   2520
      Left            =   960
      TabIndex        =   20
      Top             =   285
      Width           =   6825
      Begin VB.TextBox txt_nombre_busqueda 
         Height          =   345
         Left            =   60
         TabIndex        =   23
         Top             =   450
         Width           =   6675
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   1635
         Left            =   45
         TabIndex        =   22
         Top             =   825
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   2884
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Artículo"
            Object.Width           =   7057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Disponible"
            Object.Width           =   2470
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de cliente"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   6720
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmimportacion_clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmimportacion_clientes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7530
      Picture         =   "frmimportacion_clientes.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   15
      TabIndex        =   12
      Top             =   330
      Width           =   8070
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   105
      TabIndex        =   11
      Top             =   480
      Width           =   7830
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   330
         Left            =   2745
         TabIndex        =   7
         Top             =   1800
         Width           =   4965
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   330
         Left            =   1335
         TabIndex        =   6
         Top             =   1800
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_grupo_actual 
         Height          =   330
         Left            =   2745
         TabIndex        =   5
         Top             =   1425
         Width           =   4965
      End
      Begin VB.TextBox txt_grupo_actual 
         Height          =   330
         Left            =   1335
         TabIndex        =   4
         Top             =   1425
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_grupo_real 
         Height          =   330
         Left            =   2745
         TabIndex        =   19
         Top             =   1035
         Width           =   4965
      End
      Begin VB.TextBox txt_grupo_real 
         Height          =   330
         Left            =   1335
         TabIndex        =   18
         Top             =   1035
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   330
         Left            =   2745
         TabIndex        =   3
         Top             =   645
         Width           =   4965
      End
      Begin VB.TextBox txt_titular 
         Height          =   330
         Left            =   1335
         TabIndex        =   2
         Top             =   645
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   330
         Left            =   2745
         TabIndex        =   1
         Top             =   270
         Width           =   4965
      End
      Begin VB.TextBox txt_cliente 
         Height          =   330
         Left            =   1335
         TabIndex        =   0
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1875
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo actual:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grupo real:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1110
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   713
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   338
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmimportacion_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If Trim(Me.txt_cliente) <> "" Then
      If Trim(Me.txt_titular) <> "" Then
         If Me.txt_grupo_real <> "" Then
            If Me.txt_grupo_actual <> "" Then
               If Me.txt_establecimiento <> "rrrr" Then
                  rs.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     rsaux.Open "select * from tb_clientes WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     var_cadena = "INSERT INTO "
                     var_cadena = var_cadena + "TB_CLIENTES  (VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_CLI_REPRESENTANTE, DTIM_CLI_FECHA_CAPTURA, VCHA_AGE_AGENTE_ID, VCHA_RUT_RUTA_ID, VCHA_CLI_CURP, VCHA_CLI_RFC, VCHA_MON_MONEDA_ID, VCHA_PLA_PLAZO_ID, VCHA_TCL_TIPO_CLIENTE_ID, VCHA_LIS_LISTA_ID, VCHA_TRA_TRANSPORTE_ID, VCHA_FAG_FAMILIA_AGRUPADOR_ID, INTE_CLI_AGRUPADOR, INTE_CLI_ESTATUS, VCHA_TIT_TITULAR_ID, CHAR_PRI_PRIORIDAD_ID, VCHA_CLI_EMAIL, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_MUN_MUNICIPIO_ID, VCHA_CIU_CIUDAD_ID, VCHA_CLI_COLONIA, VCHA_CLI_DIRECCION, VCHA_CLI_CP, INTE_CLI_ENVIO_FACTURA, INTE_CLI_ASIGNACION_CATALOGOS, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, INTE_CLI_CLIENTE_PEDIDO_TIENDA, INTE_CLI_PERSONA_FISICA, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CLI_REFERENCIA, TEXTILERA, VCHA_CLI_TIENDA, VCHA_CLI_CLAVE_TIENDA, DTIM_INT_FECHA, INTE_INT_INTERFACE, INTE_CLI_FRANQUICIA, VCHA_CLI_TELEFONO, INTE_CLI_TRAZABILIDAD , Referencia, VCHA_SRU_SUBRUTA_ID,"
                     var_cadena = var_cadena + " INTE_CLI_ACTIVO, VCHA_CLI_CLAVE_UNIFICADA_ID, INTE_CLI_UNIFICADOR)  "
                     var_cadena = var_cadena + " values ('" + IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE) + "', '" + IIf(IsNull(rsaux!vcha_cli_representante), "", rsaux!vcha_cli_representante) + "', "
                     var_c = " values ('" + IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE) + "', '" + IIf(IsNull(rsaux!vcha_cli_representante), "", rsaux!vcha_cli_representante) + "', "
                     var_cadena = var_cadena + " " + Format(IIf(IsNull(rsaux!dtim_cli_fecha_Captura), Date, rsaux!dtim_cli_fecha_Captura), "Short Date") + ", '" + IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID) + "', '" + IIf(IsNull(rsaux!vcha_rut_ruta_id), "", rsaux!vcha_rut_ruta_id) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_CURP), "", rsaux!VCHA_CLI_CURP) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_RFC), "", rsaux!VCHA_CLI_RFC) + "', '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id) + "', '" + IIf(IsNull(rsaux!VCHA_PLA_PLAZO_ID), "", rsaux!VCHA_PLA_PLAZO_ID) + "', '" + IIf(IsNull(rsaux!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux!VCHA_TCL_TIPO_CLIENTE_ID) + "', '" + IIf(IsNull(rsaux!vcha_LIS_LISTA_iD), "", rsaux!vcha_LIS_LISTA_iD) + "', '" + IIf(IsNull(rsaux!VCHA_TRA_TRANSPORTE_ID), "", rsaux!VCHA_TRA_TRANSPORTE_ID) + "','"
                     var_c = var_c + " " + Format(IIf(IsNull(rsaux!dtim_cli_fecha_Captura), Date, rsaux!dtim_cli_fecha_Captura), "Short Date") + ", '" + IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID) + "', '" + IIf(IsNull(rsaux!vcha_rut_ruta_id), "", rsaux!vcha_rut_ruta_id) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_CURP), "", rsaux!VCHA_CLI_CURP) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_RFC), "", rsaux!VCHA_CLI_RFC) + "', '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id) + "', '" + IIf(IsNull(rsaux!VCHA_PLA_PLAZO_ID), "", rsaux!VCHA_PLA_PLAZO_ID) + "', '" + IIf(IsNull(rsaux!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux!VCHA_TCL_TIPO_CLIENTE_ID) + "', '" + IIf(IsNull(rsaux!vcha_LIS_LISTA_iD), "", rsaux!vcha_LIS_LISTA_iD) + "', '" + IIf(IsNull(rsaux!VCHA_TRA_TRANSPORTE_ID), "", rsaux!VCHA_TRA_TRANSPORTE_ID) + "','"
                     var_cadena = var_cadena + IIf(IsNull(rsaux!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux!VCHA_FAG_FAMILIA_AGRUPADOR_ID) + "' , " + CStr(IIf(IsNull(rsaux!INTE_CLI_AGRUPADOR), 0, rsaux!INTE_CLI_AGRUPADOR)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_ESTATUS), 0, rsaux!INTE_CLI_ESTATUS)) + ", '" + IIf(IsNull(rsaux!vcha_tit_titular_id), "", rsaux!vcha_tit_titular_id) + "', '" + IIf(IsNull(rsaux!CHAR_PRI_PRIORIDAD_ID), "", rsaux!CHAR_PRI_PRIORIDAD_ID) + "', '" + IIf(IsNull(rsaux!vcha_cli_email), "", rsaux!vcha_cli_email) + "', '"
                     var_c = var_c + IIf(IsNull(rsaux!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux!VCHA_FAG_FAMILIA_AGRUPADOR_ID) + "' , " + CStr(IIf(IsNull(rsaux!INTE_CLI_AGRUPADOR), 0, rsaux!INTE_CLI_AGRUPADOR)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_ESTATUS), 0, rsaux!INTE_CLI_ESTATUS)) + ", '" + IIf(IsNull(rsaux!vcha_tit_titular_id), "", rsaux!vcha_tit_titular_id) + "', '" + IIf(IsNull(rsaux!CHAR_PRI_PRIORIDAD_ID), "", rsaux!CHAR_PRI_PRIORIDAD_ID) + "', '" + IIf(IsNull(rsaux!vcha_cli_email), "", rsaux!vcha_cli_email) + "', '"
                     var_cadena = var_cadena + IIf(IsNull(rsaux!VCHA_PAI_PAIS_ID), "", rsaux!VCHA_PAI_PAIS_ID) + "', '" + IIf(IsNull(rsaux!VCHA_EST_ESTADO_ID), "", rsaux!VCHA_EST_ESTADO_ID) + "', '" + IIf(IsNull(rsaux!VCHA_MUN_MUNICIPIO_ID), "", rsaux!VCHA_MUN_MUNICIPIO_ID) + "', '"
                     var_c = var_c + IIf(IsNull(rsaux!VCHA_PAI_PAIS_ID), "", rsaux!VCHA_PAI_PAIS_ID) + "', '" + IIf(IsNull(rsaux!VCHA_EST_ESTADO_ID), "", rsaux!VCHA_EST_ESTADO_ID) + "', '" + IIf(IsNull(rsaux!VCHA_MUN_MUNICIPIO_ID), "", rsaux!VCHA_MUN_MUNICIPIO_ID) + "', '"
                     var_cadena = var_cadena + IIf(IsNull(rsaux!VCHA_CIU_CIUDAD_ID), "", rsaux!VCHA_CIU_CIUDAD_ID) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_COLONIA), "", rsaux!VCHA_CLI_COLONIA) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_DIRECCION), "", rsaux!VCHA_CLI_DIRECCION) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_CP), "", rsaux!VCHA_CLI_CP) + "',"
                     var_c = var_c + IIf(IsNull(rsaux!VCHA_CIU_CIUDAD_ID), "", rsaux!VCHA_CIU_CIUDAD_ID) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_COLONIA), "", rsaux!VCHA_CLI_COLONIA) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_DIRECCION), "", rsaux!VCHA_CLI_DIRECCION) + "', '" + IIf(IsNull(rsaux!VCHA_CLI_CP), "", rsaux!VCHA_CLI_CP) + "',"
                     var_cadena = var_cadena + CStr(IIf(IsNull(rsaux!INTE_CLI_ENVIO_FACTURA), 0, rsaux!INTE_CLI_ENVIO_FACTURA)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_ASIGNACION_CATALOGOS), 0, rsaux!INTE_CLI_ASIGNACION_CATALOGOS)) + ", '"
                     var_c = var_c + CStr(IIf(IsNull(rsaux!INTE_CLI_ENVIO_FACTURA), 0, rsaux!INTE_CLI_ENVIO_FACTURA)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_ASIGNACION_CATALOGOS), 0, rsaux!INTE_CLI_ASIGNACION_CATALOGOS)) + ", '"
                     var_cadena = var_cadena + IIf(IsNull(rsaux!vcha_cli_clave_anterior_id), "", rsaux!vcha_cli_clave_anterior_id) + "' , '" + IIf(IsNull(rsaux!VCHA_EMP_EMPRESA_ID), "", rsaux!VCHA_EMP_EMPRESA_ID) + "', " + CStr(IIf(IsNull(rsaux!INTE_CLI_CLIENTE_PEDIDO_TIENDA), 0, rsaux!INTE_CLI_CLIENTE_PEDIDO_TIENDA)) + ", " + CStr(IIf(IsNull(rsaux!inte_cli_persona_fisica), 0, rsaux!inte_cli_persona_fisica)) + ", " + CStr(IIf(IsNull(rsaux!num_inter_tranc_type), 0, rsaux!num_inter_tranc_type)) + ", " + CStr(IIf(IsNull(rsaux!NUM_INTER_UPLOADED), "", rsaux!NUM_INTER_UPLOADED)) + ", " + Format(IIf(IsNull(rsaux!date_inter_date), Date, rsaux!date_inter_date), "Short Date") + ", '" + IIf(IsNull(rsaux!VCHA_CLI_REFERENCIA), "", rsaux!VCHA_CLI_REFERENCIA) + "', '" + IIf(IsNull(rsaux!TEXTILERA), "", rsaux!TEXTILERA) + "', '" + IIf(IsNull(rsaux!vcha_cli_tienda), "", rsaux!vcha_cli_tienda) + "', '" + IIf(IsNull(rsaux!vcha_cli_clave_tienda), "", rsaux!vcha_cli_clave_tienda) + "', "
                     var_c = var_c + IIf(IsNull(rsaux!vcha_cli_clave_anterior_id), "", rsaux!vcha_cli_clave_anterior_id) + "' , '" + IIf(IsNull(rsaux!VCHA_EMP_EMPRESA_ID), "", rsaux!VCHA_EMP_EMPRESA_ID) + "', " + CStr(IIf(IsNull(rsaux!INTE_CLI_CLIENTE_PEDIDO_TIENDA), 0, rsaux!INTE_CLI_CLIENTE_PEDIDO_TIENDA)) + ", " + CStr(IIf(IsNull(rsaux!inte_cli_persona_fisica), 0, rsaux!inte_cli_persona_fisica)) + ", " + CStr(IIf(IsNull(rsaux!num_inter_tranc_type), 0, rsaux!num_inter_tranc_type)) + ", " + CStr(IIf(IsNull(rsaux!NUM_INTER_UPLOADED), "", rsaux!NUM_INTER_UPLOADED)) + ", " + Format(IIf(IsNull(rsaux!date_inter_date), Date, rsaux!date_inter_date), "Short Date") + ", '" + IIf(IsNull(rsaux!VCHA_CLI_REFERENCIA), "", rsaux!VCHA_CLI_REFERENCIA) + "', '" + IIf(IsNull(rsaux!TEXTILERA), "", rsaux!TEXTILERA) + "', '" + IIf(IsNull(rsaux!vcha_cli_tienda), "", rsaux!vcha_cli_tienda) + "', '" + IIf(IsNull(rsaux!vcha_cli_clave_tienda), "", rsaux!vcha_cli_clave_tienda) + "', "
                     var_cadena = var_cadena + Format(IIf(IsNull(rsaux!DTIM_INT_FECHA), Date, rsaux!DTIM_INT_FECHA), "Short Date") + ", " + CStr(IIf(IsNull(rsaux!INTE_INT_INTERFACE), 0, rsaux!INTE_INT_INTERFACE)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_FRANQUICIA), 0, rsaux!INTE_CLI_FRANQUICIA)) + ", '" + IIf(IsNull(rsaux!vcha_cli_telefono), "", rsaux!vcha_cli_telefono) + "', " + CStr(IIf(IsNull(rsaux!INTE_CLI_TRAZABILIDAD), 0, rsaux!INTE_CLI_TRAZABILIDAD)) + ", '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "', '" + IIf(IsNull(rsaux!vcha_sru_subruta_id), "", rsaux!vcha_sru_subruta_id) + "',"
                     var_c = var_c + CStr(IIf(IsNull(rsaux!DTIM_INT_FECHA), Date, rsaux!DTIM_INT_FECHA)) + ", " + CStr(IIf(IsNull(rsaux!INTE_INT_INTERFACE), 0, rsaux!INTE_INT_INTERFACE)) + ", " + CStr(IIf(IsNull(rsaux!INTE_CLI_FRANQUICIA), 0, rsaux!INTE_CLI_FRANQUICIA)) + ", '" + IIf(IsNull(rsaux!vcha_cli_telefono), "", rsaux!vcha_cli_telefono) + "', " + CStr(IIf(IsNull(rsaux!INTE_CLI_TRAZABILIDAD), 0, rsaux!INTE_CLI_TRAZABILIDAD)) + ", '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "', '" + IIf(IsNull(rsaux!vcha_sru_subruta_id), "", rsaux!vcha_sru_subruta_id) + "',"
                     var_cadena = var_cadena + CStr(IIf(IsNull(rsaux!inte_cli_activo), "", rsaux!inte_cli_activo)) + ", '" + IIf(IsNull(rsaux!vcha_cli_clave_unificada_id), "", rsaux!vcha_cli_clave_unificada_id) + "', " + CStr(IIf(IsNull(rsaux!inte_cli_unificador), 0, rsaux!inte_cli_unificador)) + ")"
                     var_c = var_c + CStr(IIf(IsNull(rsaux!inte_cli_activo), "", rsaux!inte_cli_activo)) + ", '" + IIf(IsNull(rsaux!vcha_cli_clave_unificada_id), "", rsaux!vcha_cli_clave_unificada_id) + "', " + CStr(IIf(IsNull(rsaux!inte_cli_unificador), 0, rsaux!inte_cli_unificador)) + ")"
                     MsgBox var_c
                     rsaux.Close
                     rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux2.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     var_cadena = "UPDATE TB_CLIENTES SET VCHA_CLI_NOMBRE = '" + rsaux2!VCHA_CLI_NOMBRE + "', VCHA_CLI_REPRESENTANTE = '" + IIf(IsNull(rsaux2!vcha_cli_representante), "", rsaux2!vcha_cli_representante) + "', DTIM_CLI_FECHA_CAPTURA = " + CStr(IIf(IsNull(rsaux2!dtim_cli_fecha_Captura), "", rsaux2!dtim_cli_fecha_Captura)) + ", VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rsaux2!VCHA_AGE_AGENTE_ID), "", rsaux2!VCHA_AGE_AGENTE_ID) + "', VCHA_RUT_RUTA_ID = '" + IIf(IsNull(rsaux2!vcha_rut_ruta_id), "", rsaux2!vcha_rut_ruta_id) + "', VCHA_CLI_CURP = '" + IIf(IsNull(rsaux2!VCHA_CLI_CURP), "", rsaux2!VCHA_CLI_CURP) + "', VCHA_CLI_RFC = '" + IIf(IsNull(rsaux2!VCHA_CLI_RFC), "", rsaux2!VCHA_CLI_RFC) + "', VCHA_MON_MONEDA_ID = '" + IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id) + "', VCHA_PLA_PLAZO_ID = '" + IIf(IsNull(rsaux2!VCHA_PLA_PLAZO_ID), "", rsaux2!VCHA_PLA_PLAZO_ID) + "', VCHA_TCL_TIPO_CLIENTE_ID = '"
                     var_cadena = var_cadena + IIf(IsNull(rsaux2!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux2!VCHA_TCL_TIPO_CLIENTE_ID) + "',VCHA_LIS_LISTA_ID = '" + IIf(IsNull(rsaux2!vcha_LIS_LISTA_iD), "", rsaux2!vcha_LIS_LISTA_iD) + "', VCHA_TRA_TRANSPORTE_ID = '" + IIf(IsNull(rsaux2!VCHA_TRA_TRANSPORTE_ID), "", rsaux2!VCHA_TRA_TRANSPORTE_ID) + "', VCHA_FAG_FAMILIA_AGRUPADOR_ID = '" + IIf(IsNull(rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID), "", rsaux2!VCHA_FAG_FAMILIA_AGRUPADOR_ID) + "', INTE_CLI_AGRUPADOR = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_AGRUPADOR), 0, rsaux2!INTE_CLI_AGRUPADOR)) + ", INTE_CLI_ESTATUS = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_ESTATUS), 0, rsaux2!INTE_CLI_ESTATUS)) + ", VCHA_TIT_TITULAR_ID = '" + IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id) + "', CHAR_PRI_PRIORIDAD_ID = '" + IIf(IsNull(rsaux2!CHAR_PRI_PRIORIDAD_ID), "", rsaux2!CHAR_PRI_PRIORIDAD_ID) + "', VCHA_CLI_EMAIL = '" + IIf(IsNull(rsaux2!vcha_cli_email), "", rsaux2!vcha_cli_email) + "', "
                     var_cadena = var_cadena + " VCHA_PAI_PAIS_ID = '" + IIf(IsNull(rsaux2!VCHA_PAI_PAIS_ID), "", rsaux2!VCHA_PAI_PAIS_ID) + "', VCHA_EST_ESTADO_ID = '" + IIf(IsNull(rsaux2!VCHA_EST_ESTADO_ID), "", rsaux2!VCHA_EST_ESTADO_ID) + "', VCHA_MUN_MUNICIPIO_ID = '" + IIf(IsNull(rsaux2!VCHA_MUN_MUNICIPIO_ID), "", rsaux2!VCHA_MUN_MUNICIPIO_ID) + "', VCHA_CIU_CIUDAD_ID = '" + IIf(IsNull(rsaux2!VCHA_CIU_CIUDAD_ID), "", rsaux2!VCHA_CIU_CIUDAD_ID) + "', VCHA_cLI_COLONIA = '" + IIf(IsNull(rsaux2!VCHA_CLI_COLONIA), "", rsaux2!VCHA_CLI_COLONIA) + "', VCHA_CLI_DIRECCION = '" + IIf(IsNull(rsaux2!VCHA_CLI_DIRECCION), "", rsaux2!VCHA_CLI_DIRECCION) + "', VCHA_CLI_CP ='" + IIf(IsNull(rsaux2!VCHA_CLI_CP), "", rsaux2!VCHA_CLI_CP) + "',INTE_CLI_ENVIO_FACTURA= " + CStr(IIf(IsNull(rsaux2!INTE_CLI_ENVIO_FACTURA), 0, rsaux2!INTE_CLI_ENVIO_FACTURA)) + ", INTE_CLI_ASIGNACION_CATALOGOS = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_ASIGNACION_CATALOGOS), 0, rsaux2!INTE_CLI_ASIGNACION_CATALOGOS)) + ", VCHA_CLI_CLAVE_ANTERIOR_ID = '"
                     var_cadena = var_cadena + IIf(IsNull(rsaux2!vcha_cli_clave_anterior_id), "", rsaux2!vcha_cli_clave_anterior_id) + "  ', VCHA_EMP_EMPRESA_ID = '" + IIf(IsNull(rsaux2!VCHA_EMP_EMPRESA_ID), "", rsaux2!VCHA_EMP_EMPRESA_ID) + "', INTE_CLI_CLIENTE_PEDIDO_TIENDA = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_CLIENTE_PEDIDO_TIENDA), 0, rsaux2!INTE_CLI_CLIENTE_PEDIDO_TIENDA)) + ", INTE_CLI_PERSONA_FISICA = " + CStr(IIf(IsNull(rsaux2!inte_cli_persona_fisica), 0, rsaux2!inte_cli_persona_fisica)) + ", TEXTILERA = '" + IIf(IsNull(rsaux2!TEXTILERA), "", rsaux2!TEXTILERA) + "', VCHA_CLI_REFERENCIA = '" + IIf(IsNull(rsaux2!VCHA_CLI_REFERENCIA), "", rsaux2!VCHA_CLI_REFERENCIA) + "', VCHA_CLI_TIENDA = '" + IIf(IsNull(rsaux2!vcha_cli_tienda), "", rsaux2!vcha_cli_tienda) + "', VCHA_CLI_CLAVE_TIENDA = '" + IIf(IsNull(rsaux2!vcha_cli_clave_tienda), "", rsaux2!vcha_cli_clave_tienda) + "', INTE_CLI_FRANQUICIA = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_FRANQUICIA), 0, rsaux2!INTE_CLI_FRANQUICIA)) + ",  "
                     var_cadena = var_cadena + " VCHA_CLI_TELEFONO = '" + IIf(IsNull(rsaux2!vcha_cli_telefono), "", rsaux2!vcha_cli_telefono) + "', INTE_CLI_TRAZABILIDAD = " + CStr(IIf(IsNull(rsaux2!INTE_CLI_TRAZABILIDAD), 0, rsaux2!INTE_CLI_TRAZABILIDAD)) + " WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'"
                     Text1 = var_cadena
                     rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux2.Close
                  End If
                  rs.Close
                  rs.Open "select * from tb_titulares where vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     rsaux.Open "insert into " + parametros(0) + "." + parametros(1) + ".dbo.tb_titulares (VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_MUN_MUNICIPIO_ID, VCHA_CIU_CIUDAD_ID,  VCHA_COL_COLONIA_ID, VCHA_TIT_DOMICILIO, VCHA_TIT_CP, VCHA_TIT_TELEFONO, FLOA_TIT_LIMITE_CREDITO, VCHA_TIT_TITULAR_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, TEXTILERA, DTIM_INT_FECHA, INTE_INT_INTERFACE, VCHA_TIT_EMAIL) select VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_MUN_MUNICIPIO_ID, VCHA_CIU_CIUDAD_ID,  VCHA_COL_COLONIA_ID, VCHA_TIT_DOMICILIO, VCHA_TIT_CP, VCHA_TIT_TELEFONO, FLOA_TIT_LIMITE_CREDITO, VCHA_TIT_TITULAR_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, TEXTILERA, DTIM_INT_FECHA, INTE_INT_INTERFACE, VCHA_TIT_EMAIL from tb_titulares where vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux2.Open "select * from tb_titulares where vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     var_cadena = "update tb_titulares set vcha_gre_grupo_real_id = '" + IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id) + "', vcha_tit_nombre = '" + IIf(IsNull(rsaux2!VCHA_TIT_NOMBRE), "", rsaux2!VCHA_TIT_NOMBRE) + "', vcha_pai_pais_id = '" + IIf(IsNull(rsaux2!VCHA_PAI_PAIS_ID), "", rsaux2!VCHA_PAI_PAIS_ID) + "', vcha_est_Estado_id = '" + IIf(IsNull(rsaux2!VCHA_EST_ESTADO_ID), "", rsaux2!VCHA_EST_ESTADO_ID) + "', vcha_mun_municipio_id = '" + IIf(IsNull(rsaux2!VCHA_MUN_MUNICIPIO_ID), "", rsaux2!VCHA_MUN_MUNICIPIO_ID) + "', vcha_ciu_ciudad_id = '" + IIf(IsNull(rsaux2!VCHA_CIU_CIUDAD_ID), "", rsaux2!VCHA_CIU_CIUDAD_ID) + "', vcha_col_colonia_id = '" + IIf(IsNull(rsaux2!VCHA_COL_COLONIA_ID), "", rsaux2!VCHA_COL_COLONIA_ID) + "', vcha_tit_domicilio = '" + IIf(IsNull(rsaux2!VCHA_TIT_DOMICILIO), "", rsaux2!VCHA_TIT_DOMICILIO) + "', vcha_tit_cp = '" + IIf(IsNull(rsaux2!VCHA_TIT_CP), "", rsaux2!VCHA_TIT_CP) + "', "
                     var_cadena = var_cadena + " vcha_tit_telefono = '" + IIf(IsNull(rsaux2!VCHA_TIT_TELEFONO), "", rsaux2!VCHA_TIT_TELEFONO) + "', floa_tit_limite_credito = " + CStr(IIf(IsNull(rsaux2!floa_tit_limite_credito), 0, rsaux2!floa_tit_limite_credito)) + ", vcha_tit_titular_anterior_id = '" + IIf(IsNull(rsaux2!VCHA_TIT_TITULAR_ANTERIOR_ID), "", rsaux2!vcha_tit_titular_id) + "', vcha_emp_empresa_id = '" + IIf(IsNull(rsaux2!VCHA_EMP_EMPRESA_ID), "", rsaux2!VCHA_EMP_EMPRESA_ID) + "', textilera = '" + IIf(IsNull(rsaux2!TEXTILERA), "", rsaux2!TEXTILERA) + "', vcha_tit_email = '" + IIf(IsNull(rsaux2!VCHA_TIT_EMAIL), "", rsaux2!VCHA_TIT_EMAIL) + "' where vcha_tit_titular_id = '" + Me.txt_titular + "'"
                     'MsgBox var_cadena
                     rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux2.Close
                  End If
                  rs.Close
                  rs.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + Me.txt_grupo_real + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     rsaux.Open "insert into " + parametros(0) + "." + parametros(1) + ".dbo.tb_gruposreales (VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2, FLOA_GRE_DESCUENTO_3, CHAR_PRI_PRIORIDAD_ID, VCHA_GRE_GRUPO_REAL_ANTERIOR_ID, VCHA_GAC_GRUPO_ACTUAL_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, TEXTILERA, DTIM_INT_FECHA, INTE_INT_INTERFACE) select VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_GRE_NOMBRE, FLOA_GRE_DESCUENTO_1, FLOA_GRE_DESCUENTO_2, FLOA_GRE_DESCUENTO_3, CHAR_PRI_PRIORIDAD_ID, VCHA_GRE_GRUPO_REAL_ANTERIOR_ID, VCHA_GAC_GRUPO_ACTUAL_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, TEXTILERA, DTIM_INT_FECHA, INTE_INT_INTERFACE from DISTRIBUCION.VIANNEY.DBO.tb_gruposreales where vcha_gre_grupo_real_id = '" + Me.txt_grupo_real + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "update tb_gruposreales set vcha_gre_nombre= '" + Me.txt_nombre_grupo_real + "' where vcha_gre_grupo_real_id = '" + Me.txt_grupo_real + "'", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
                  rs.Open "select * from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + Me.txt_grupo_actual + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     rsaux.Open "insert into " + parametros(0) + "." + parametros(1) + ".dbo.tb_gruposactuales (vcha_gac_grupo_actual_id, vcha_gac_nombre, floa_gac_Descuento_1, floa_gac_descuento_2, floa_gac_descuento_3, vcha_gac_grupo_actual_anterior_id, vcha_emp_empresa_id, textilera, inte_gac_tela) select vcha_gac_grupo_actual_id, vcha_gac_nombre, floa_gac_Descuento_1, floa_gac_descuento_2, floa_gac_descuento_3, vcha_gac_grupo_actual_anterior_id, vcha_emp_empresa_id, textilera, inte_gac_tela from distribucion.vianney.dbo.tb_gruposactuales where vcha_gac_grupo_Actual_id = '" + Me.txt_grupo_actual + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "UPDATE tb_gruposactuales set vcha_gac_nombre = '" + Me.txt_nombre_grupo_actual + "' where vcha_gac_grupo_Actual_id = '" + Me.txt_grupo_actual + "'", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
                  If Me.txt_establecimiento <> "" Then
                     rs.Open "select * from tb_establecimientos where vcha_Esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rs.EOF Then
                        rsaux.Open "insert into tb_establecimientos (VCHA_TIT_TITULAR_ID, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_CIU_CIUDAD_ID, VCHA_COL_COLONIA_ID, VCHA_ESB_DOMICILIO, VCHA_ESB_TELEFONO, CHAR_ESB_FACTURA_CATALOGOS, VCHA_MUN_MUNICIPIO_ID, VCHA_ESB_CP, VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, INTE_ESB_FRANQUICIA) select VCHA_TIT_TITULAR_ID, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_CIU_CIUDAD_ID, VCHA_COL_COLONIA_ID, VCHA_ESB_DOMICILIO, VCHA_ESB_TELEFONO, CHAR_ESB_FACTURA_CATALOGOS, VCHA_MUN_MUNICIPIO_ID, VCHA_ESB_CP, VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID, VCHA_EMP_EMPRESA_ID, INTE_ESB_FRANQUICIA from distribucion.vianney.dbo.tb_Establecimientos where vcha_esb_Establecimiento_id = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux.Open "update tb_establecimientos set vcha_esb_nombre = '" + Me.txt_nombre_establecimiento + "' where vcha_esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rs.Close
                     rs.Open "select * from tb_detalle_Establecimientos where vcha_esb_establecimiento_id = '" + Me.txt_establecimiento + "' and vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rs.EOF Then
                        rsaux.Open "insert into tb_Detalle_establecimientos (vcha_Esb_establecimiento_id, vcha_cli_clave_id) values ('" + Me.txt_establecimiento + "','" + Me.txt_cliente + "')", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rs.Close
                  End If
                  MsgBox "Se a terminado de actualizar los datos del cliente", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El cliente seleccionado no cuenta con un establecimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El cliente seleccionado no cuenta con un grupo actual", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El cliente seleccionado no cuenta con un grupo real", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El cliente seleccionado no cuenta con titular", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Top = 2000
   Me.Left = 2000
   Me.frm_busqueda_folio.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cliente = Me.lv_articulos.selectedItem
      
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda_folio.Visible = False
   End If
End Sub

Private Sub txt_cliente_Change()
   Me.txt_establecimiento = ""
   Me.txt_grupo_actual = ""
   Me.txt_grupo_real = ""
   Me.txt_nombre_cliente = ""
   Me.txt_nombre_establecimiento = ""
   Me.txt_nombre_grupo_actual = ""
   Me.txt_nombre_grupo_real = ""
   Me.txt_nombre_titular = ""
   Me.txt_titular = ""
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(Me.txt_cliente) <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      'MsgBox cnn_distribucion.ConnectionString
      rs.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.txt_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         Me.txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         Me.txt_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         Me.txt_nombre_grupo_real = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
         Me.txt_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
         Me.txt_nombre_grupo_actual = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
         rsaux.Open "SELECT * FROM TB_DETALLE_ESTABLECIMIENTOS WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id = '" + rsaux!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            Me.txt_establecimiento = IIf(IsNull(rsaux1!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux1!vcha_ESB_ESTABLECIMIENTO_id)
            Me.txt_nombre_establecimiento = IIf(IsNull(rsaux1!VCHA_ESB_NOMBRE), "", rsaux1!VCHA_ESB_NOMBRE)
            rsaux1.Close
         Else
            MsgBox "El cliente no tiene un establecimiento asociado", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      Else
         MsgBox "La clave del cliente no existe", vbOKOnly, "ATENCION"
         Me.txt_establecimiento = ""
         Me.txt_grupo_actual = ""
         Me.txt_grupo_real = ""
         Me.txt_nombre_cliente = ""
         Me.txt_nombre_establecimiento = ""
         Me.txt_nombre_grupo_actual = ""
         Me.txt_nombre_grupo_real = ""
         Me.txt_nombre_titular = ""
         Me.txt_titular = ""
      End If
      rs.Close
   Else
      Me.txt_establecimiento = ""
      Me.txt_grupo_actual = ""
      Me.txt_grupo_real = ""
      Me.txt_nombre_cliente = ""
      Me.txt_nombre_establecimiento = ""
      Me.txt_nombre_grupo_actual = ""
      Me.txt_nombre_grupo_real = ""
      Me.txt_nombre_titular = ""
      Me.txt_titular = ""
   End If
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_grupo_real_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_articulo_Change()

End Sub


Private Sub txt_nombre_busqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_busqueda_folio.Visible = False
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from vw_clientes where vcha_cli_nombre like '%" + Me.txt_nombre_busqueda + "%' and vcha_emp_Empresa_id = '" + var_empresa + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      rs.Close
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Me.cmd_guardar.SetFocus
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_grupo_real_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      Else
         Unload Me
      End If
   Else
      KeyAscii = 0
   End If
End Sub
