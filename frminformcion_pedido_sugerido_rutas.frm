VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frminformacion_pedido_sugerido_rutas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información para pedido sugerido"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView mon_mes2 
      Height          =   2370
      Left            =   2940
      TabIndex        =   7
      Top             =   660
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   72876033
      CurrentDate     =   37761
   End
   Begin MSComCtl2.MonthView mon_mes1 
      Height          =   2370
      Left            =   300
      TabIndex        =   6
      Top             =   660
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   72876033
      CurrentDate     =   37761
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   45
      TabIndex        =   20
      Top             =   570
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   21
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
         TabIndex        =   22
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   17
      Top             =   390
      Width           =   5805
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Datos "
      Height          =   1695
      Left            =   75
      TabIndex        =   16
      Top             =   510
      Width           =   5625
      Begin VB.TextBox txt_nombre_ruta 
         Height          =   345
         Left            =   1245
         TabIndex        =   3
         Top             =   1215
         Width           =   4275
      End
      Begin VB.TextBox txt_ruta 
         Height          =   345
         Left            =   135
         TabIndex        =   2
         Top             =   1215
         Width           =   1065
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   345
         Left            =   1245
         TabIndex        =   1
         Top             =   510
         Width           =   4275
      End
      Begin VB.TextBox txt_agente 
         Height          =   345
         Left            =   135
         TabIndex        =   0
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   975
         Width           =   345
      End
      Begin VB.Label Label3 
         Caption         =   "Agente"
         Height          =   240
         Left            =   165
         TabIndex        =   18
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5370
      Picture         =   "frminformcion_pedido_sugerido_rutas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frminformcion_pedido_sugerido_rutas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Enviar Información"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frminformcion_pedido_sugerido_rutas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   75
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Periodo "
      Height          =   765
      Left            =   75
      TabIndex        =   8
      Top             =   2250
      Width           =   5625
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   855
         TabIndex        =   4
         Top             =   315
         Width           =   1200
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         Top             =   315
         Width           =   1200
      End
      Begin VB.CommandButton cmd_mes_1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2085
         Picture         =   "frminformcion_pedido_sugerido_rutas.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Nuevo Movimiento Alt + N"
         Top             =   315
         Width           =   330
      End
      Begin VB.CommandButton cmd_mes_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5115
         Picture         =   "frminformcion_pedido_sugerido_rutas.frx":1AB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Nuevo Movimiento Alt + N"
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3525
         TabIndex        =   11
         Top             =   375
         Width           =   255
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2445
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1515
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frminformacion_pedido_sugerido_rutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer

Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function



Private Sub cmb_agentes_Click()
   txt_agente = Obtener_llave(cnn, rs, "tb_rutas", "VCHA_rut_NOMBRE", cmb_agentes, 0, "T")
End Sub

Private Sub cmb_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fecha_inicio.SetFocus
   End If
End Sub

Private Sub cmd_correo_Click()
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   cnn.CommandTimeout = 360
   'rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   'var_ruta = rs!VCHA_PRI_RUTA_PEDIDO_SUGERIDO
   'rs.Close
   rs.Open "select * from tb_Agentes where vcha_age_agente_id= '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_ruta = IIf(IsNull(rs!VCHA_AGE_RUTA_ARCHIVOS), "", rs!VCHA_AGE_RUTA_ARCHIVOS)
   End If
   rs.Close
   If Trim(var_ruta) <> "" Then
      var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   End If
   If Trim(txt_agente) <> "" Then
      If Trim(txt_ruta) <> "" Then
         If IsDate(txt_fecha_inicio) Then
            If IsDate(txt_fecha_fin) Then
               Dim var_especial As String
            Dim var_Archivo As String
            Dim var_tipo_agrupamiento As String
            Dim var_fecha_fin As Date
            Dim var_fecha As Date
            Dim var_correo_electronico As String
            var_fecha_fin = txt_fecha_fin
            rs.Open "select CHAR_PRI_TIPO_AGRUPAMIENTO from tb_principal", cnn, adOpenDynamic, adLockOptimistic
            var_tipo_agrupamiento = rs(0).Value
            rs.Close
            rs.Open "SELECT VCHA_AGE_AGENTE_ANTERIOR_ID FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            VAR_CLAVE_AGENTE_ANTERIOR = IIf(IsNull(rs!VCHA_AGE_AGENTE_ANTERIOR_ID), "", rs!VCHA_AGE_AGENTE_ANTERIOR_ID)
            rs.Close
            rs.Open "SELECT vcha_rut_ruta_anterior_id FROM tb_rutas WHERE vcha_rut_ruta_id = '" + txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
            VAR_CLAVE_AGENTE_ANTERIOR = IIf(IsNull(rs!VCHA_rut_ruta_ANTERIOR_ID), "", rs!VCHA_rut_ruta_ANTERIOR_ID)
            rs.Close
            dia_S = CStr(Day(txt_fecha_fin))
            If Len(Trim(dia_S)) = 1 Then
               dia_S = "0" + dia_S
            End If
            MES_S = CStr(Month(txt_fecha_fin))
            If Len(Trim(MES_S)) = 1 Then
               MES_S = "0" + MES_S
            End If
            var_Archivo = Trim(VAR_CLAVE_AGENTE_ANTERIOR) + Right(CStr(Year(txt_fecha_fin)), 2) + MES_S + dia_S
            var_eliminar = DeleteFile(var_ruta & "tem_clientes.dbf")
            var_eliminar = DeleteFile(var_ruta & "clientes.dbf")
            var_copia = CopyFile(var_ruta & "tclientes.dbf", var_ruta & "tem_clientes.dbf", 1)
            var_eliminar = DeleteFile(var_ruta & "tem_detatien.dbf")
            var_eliminar = DeleteFile(var_ruta & "detatien.dbf")
            var_copia = CopyFile(var_ruta & "tdetatien.dbf", var_ruta & "tem_detatien.dbf", 1)
            var_eliminar = DeleteFile(var_ruta & "tem_tiendas.dbf")
            var_eliminar = DeleteFile(var_ruta & "tiendas.dbf")
            var_copia = CopyFile(var_ruta & "ttiendas.dbf", var_ruta & "tem_tiendas.dbf", 1)
            var_eliminar = DeleteFile(var_ruta & "tem_titular.dbf")
            var_eliminar = DeleteFile(var_ruta & "titular.dbf")
            var_copia = CopyFile(var_ruta & "ttitular.dbf", var_ruta & "tem_titular.dbf", 1)
            var_eliminar = DeleteFile(var_ruta & "tem_" + var_Archivo + ".dbf")
            var_eliminar = DeleteFile(var_ruta & var_Archivo + ".dbf")
            var_copia = CopyFile(var_ruta & "facturas.dbf", var_ruta + "tem_" + var_Archivo + ".dbf", 1)
            'MsgBox var_tabla.ConnectionString
            rs.Open "DELETE FROM tem_" + var_Archivo, var_tabla, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM tem_clientes", var_tabla, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM tem_detatien", var_tabla, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM tem_tiendas", var_tabla, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM tem_titular", var_tabla, adOpenDynamic, adLockOptimistic
            
            
            rs.Open "select * from vw_titulares_1 where vcha_age_agente_id = '" + txt_agente + "' and vcha_rut_ruta_id = '" + txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
            
            var_cadena = "insert into tem_" + var_Archivo + ".dbf (cveestilo, cvecliente, cvetienda, fecha, especial, cantpedi1, cantpedi2, cantpedi3, cantpedi4, cantpedi5, cantpedi6, cantsurt1, cantsurt2, cantsurt3, cantsurt4, cantsurt5, cantsurt6, importe, dcto, finiperiod, ffinperiod) values "
            var_cadena = var_cadena + " ('PERIODO', 'PERIODO', 'PERIODO', ctod('" + Format(CDate(Me.txt_fecha_fin), "mm/dd/yy") + "'), 0,0, "
            var_cadena = var_cadena + "0,0,0,0,0,0, 0,0,0,0,0,0,'P', ctod('" + Format(CDate(Me.txt_fecha_inicio), "mm/dd/yy") + "'), ctod('" + Format(CDate(Me.txt_fecha_fin), "mm/dd/yy") + "'))"
            rsaux4.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into tem_titular (cvetitular, nombre, rfc, direccion, colonia, ciudad, cveciudad, cveestado, telefono, cvepais, codigopost, fechaalta, status) values "
                     var_cadena = var_cadena + "('" + IIf(IsNull(rs!VCHA_TIT_TITULAR_ANTERIOR_ID), "", rs!VCHA_TIT_TITULAR_ANTERIOR_ID) + "', '" + IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE) + "', '', '" + IIf(IsNull(rs!VCHA_TIT_DOMICILIO), "", rs!VCHA_TIT_DOMICILIO) + "', '" + IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID) + "', '" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + "', '"
                     var_cadena = var_cadena + IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID) + "', '" + IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID) + "', '" + IIf(IsNull(rs!VCHA_TIT_TELEFONO), "", rs!VCHA_TIT_TELEFONO) + "', '" + IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID) + "', '',ctod('" + CStr(Date) + "'),'')"
                     rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
            End If
            rs.Close
            rs.Open "select * from vw_establecimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' and vcha_rut_ruta_id = '" + txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     rsaux.Open "insert into tem_detatien (cvecliente, cvetienda) values ('" + IIf(IsNull(rs!vcha_cli_clave_anterior_id), "", rs!vcha_cli_clave_anterior_id) + "','" + IIf(IsNull(rs!vcha_esb_establecimiento_anterior_id), "", rs!vcha_esb_establecimiento_anterior_id) + "')", var_tabla, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
            End If
            rs.Close
            rs.Open "select distinct * from vw_establecimientos_2 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' and vcha_rut_ruta_id = '" + txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into tem_tiendas (cvetitular, cvetienda, direccion, colonia, ciudad, cveciudad, cveestado, cvepais, telefono, codigopost, fechaalta, status) values "
                     var_cadena = var_cadena + "('" + IIf(IsNull(rs!VCHA_TIT_TITULAR_ANTERIOR_ID), "", rs!VCHA_TIT_TITULAR_ANTERIOR_ID) + "', '" + IIf(IsNull(rs!vcha_esb_establecimiento_anterior_id), "", rs!vcha_esb_establecimiento_anterior_id) + "',  '" + IIf(IsNull(rs!vcha_esb_domicilio), "", rs!vcha_esb_domicilio) + "', '" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + "', '" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + "', '"
                     var_cadena = var_cadena + IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID) + "', '" + IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID) + "', '" + IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID) + "', '" + IIf(IsNull(rs!vcha_esb_telefono), "", rs!vcha_esb_telefono) + "', '',ctod('" + CStr(Date) + "'),'')"
                     rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
            End If
            rs.Close
            Text1 = "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' and len(vcha_tit_titular_id) = 10 and vcha_rut_ruta_id = '" + txt_ruta + "' order by vcha_cli_nombre"
            rs.Open "select * from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' and len(vcha_tit_titular_id) = 10 and vcha_rut_ruta_id = '" + txt_ruta + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into tem_clientes (cveempresa, cvecliente, razonsocia, direccion, colonia, rfc, telefono, ciudadold, ciudad, cveciudad, estado, codigopost, pais, zona, cveagente, fechaalta, limitecred, descuento, dctoptopag, plazopagar, diaspronpa, fecultcomp, feultabono, saldinicio, caracumano, aboacumano, grupo, cte, tipoclient, codigo, cadena, status, atencion, diavta, diarev, horarev, diacob, horacob, diamto, tprecio, activo, cvetitular, cvefamcte, cvefamdcto,periodo,referencia,agente,ruta) values "
                     var_cadena = var_cadena + "('" + IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID) + "', '" + IIf(IsNull(rs!vcha_cli_clave_anterior_id), "", rs!vcha_cli_clave_anterior_id) + "', '" + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE) + "', '" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + "', '" + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA) + "', '"
                     var_cadena = var_cadena + IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC) + "', '" + IIf(IsNull(rs!VCHA_TIT_TELEFONO), "", rs!VCHA_TIT_TELEFONO) + "', '" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + "', '" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + "', '" + IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID) + "', '" + IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID) + "', '"
                     var_cadena = var_cadena + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + "', '" + IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID) + "', '', '" + IIf(IsNull(rs!VCHA_rut_ruta_ANTERIOR_ID), "", rs!VCHA_rut_ruta_ANTERIOR_ID) + "', ctod('" + CStr(IIf(IsNull(rs!dtim_cli_fecha_Captura), "", rs!dtim_cli_fecha_Captura)) + "'), " + CStr(IIf(IsNull(rs!floa_tit_limite_credito), 0, rs!floa_tit_limite_credito)) + ", "
                     If var_tipo_agrupamiento = "A" Then
                        var_cadena = var_cadena + CStr(IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)) + ", " + CStr(IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)) + ", " + CStr(IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias))
                     End If
                     If var_tipo_agrupamiento = "R" Then
                        var_cadena = var_cadena + CStr(IIf(IsNull(rs!floa_gre_descuento_1), 0, rs!floa_gre_descuento_1)) + ", " + CStr(IIf(IsNull(rs!floa_gre_descuento_2), 0, rs!floa_gre_descuento_2)) + ", " + CStr(IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias))
                     End If
                     var_cadena = var_cadena + ",0,ctod(''), ctod(''), 0, 0, 0, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 0, '" + IIf(IsNull(rs!VCHA_TIT_TITULAR_ANTERIOR_ID), "", rs!VCHA_TIT_TITULAR_ANTERIOR_ID) + "', '" + IIf(IsNull(rs!vcha_gre_grupo_real_anterior_id), "", rs!vcha_gre_grupo_real_anterior_id) + "', '" + IIf(IsNull(rs!vcha_gac_grupo_actual_Anterior_id), "", rs!vcha_gac_grupo_actual_Anterior_id) + "','','" + IIf(IsNull(rs!VCHA_CLI_REFERENCIA), "", rs!VCHA_CLI_REFERENCIA) + "'," + CStr(CDbl(IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "0", rs!VCHA_AGE_AGENTE_ID))) + ",'" + IIf(IsNull(rs!vcha_rut_ruta_id), "", rs!vcha_rut_ruta_id) + "')"
                     rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
            End If
            rs.Close
            'var_cadena = "SELECT SUM(dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.CANTIDAD) AS CANTIDAD_PEDIDA, SUM(dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.CANTIDAD_FACTURADA) AS cantidad_facturada,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_EMP_EMPRESA_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_UOR_UNIDAD_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_AGE_AGENTE_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_TIT_TITULAR_ANTERIOR_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_CLI_CLAVE_ANTERIOR_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.INTE_PED_ESPECIALES, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.INTE_PED_SUGERIDO,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_ART_CODIGO_EXTERNO as articulo, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.AÑO,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.mes , dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.dia"
            'var_cadena = var_cadena + " FROM dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION INNER JOIN"
            'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS ON"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID AND"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID AND"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO"
            'var_cadena = var_cadena + " WHERE (dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA BETWEEN '" + txt_fecha_inicio + "' AND '" + CStr(var_fecha_fin + 1) + "')"
            'var_cadena = var_cadena + " and dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_AGE_AGENTE_ID = '" + txt_agente + "' "
            'var_cadena = var_cadena + " and dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' "
            'var_cadena = var_cadena + " GROUP BY dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_EMP_EMPRESA_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_UOR_UNIDAD_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_AGE_AGENTE_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_TIT_TITULAR_ID,dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_TIT_TITULAR_ANTERIOR_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_CLI_CLAVE_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_CLI_CLAVE_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_ESB_ESTABLECIMIENTO_ID,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.INTE_PED_ESPECIALES, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.INTE_PED_SUGERIDO,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_ART_CODIGO_EXTERNO,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.AÑO, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.MES,"
            'var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.dia, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_UNION.VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_TIT_TITULAR_ANTERIOR_ID, VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID"
            
            
             
             
             
            var_cadena = "SELECT dbo.TB_DETALLE_PEDIDOS.VCHA_ART_ARTICULO_ID, SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS pedida, SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD_SURTIDA) AS cantidad_surtida, dbo.TB_ARTICULOS.VCHA_ART_CODIGO_EXTERNO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ANTERIOR_ID, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO , dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID, day(dbo.TB_ENCABEZADO_PEDIDOS.dtim_ped_fecha) as dia, month(dbo.TB_ENCABEZADO_PEDIDOS.dtim_ped_fecha) as mes, year(dbo.TB_ENCABEZADO_PEDIDOS.dtim_ped_fecha) as año "
            var_cadena = var_cadena + " FROM dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DETALLE_PEDIDOS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_ESTABLECIMIENTOS ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID WHERE (dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA BETWEEN '" + txt_fecha_inicio + "' AND '" + CStr(var_fecha_fin + 1) + "')"
            var_cadena = var_cadena + " and (dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = '" + txt_agente + "') and (dbo.tb_clientes.vcha_rut_ruta_id = '" + txt_ruta + "')  gROUP BY dbo.TB_DETALLE_PEDIDOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_CODIGO_EXTERNO, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ANTERIOR_ID, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO , dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID"
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  If IsNull(rs!inte_ped_sugerido) Then
                     var_especial = ".F."
                  Else
                     If rs!inte_ped_sugerido = 0 Then
                        var_especial = ".F."
                     Else
                        var_especial = ".T."
                     End If
                  End If
                  var_fecha = Format(CStr(rs!mes) + "/" + CStr(rs!dia) + "/" + CStr(rs!año), "mm/dd/yy")
                  var_cadena = "insert into tem_" + var_Archivo + ".dbf (cveestilo, cvecliente, cvetienda, fecha, especial, cantpedi1, cantpedi2, cantpedi3, cantpedi4, cantpedi5, cantpedi6, cantsurt1, cantsurt2, cantsurt3, cantsurt4, cantsurt5, cantsurt6, importe, dcto, finiperiod, ffinperiod) values "
                  var_cadena = var_cadena + " ('" + IIf(IsNull(rs!VCHA_aRT_CODIGO_EXTERNO), "", rs!VCHA_aRT_CODIGO_EXTERNO) + "', '" + IIf(IsNull(rs!vcha_cli_clave_anterior_id), "", rs!vcha_cli_clave_anterior_id) + "', '" + IIf(IsNull(rs!vcha_esb_establecimiento_anterior_id), "", rs!vcha_esb_establecimiento_anterior_id) + "', CTOD('" + CStr(rs!mes) + "/" + CStr(rs!dia) + "/" + CStr(rs!año) + "'), " + var_especial + "," + CStr(IIf(IsNull(rs!pedida), 0, rs!pedida)) + ", "
                  var_cadena = var_cadena + "0,0,0,0,0," + CStr(IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida)) + ", 0,0,0,0,0,0,'M',ctod('" + Format(CDate(Me.txt_fecha_inicio), "mm/dd/yy") + "'), ctod('" + Format(CDate(Me.txt_fecha_fin), "mm/dd/yy") + "'))"
                  rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
               Wend
            End If
            rs.Close
            var_cadena = "SELECT  SUM(dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.CANTIDAD_DEVUELTA) AS cantidad_devuelta, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_ART_CODIGO_EXTERNO, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_AGE_AGENTE_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID,"
            var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_CLI_CLAVE_aNTERIOR_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.AÑO, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.mes , dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.dia FROM dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND"
            var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_CAR_TIPO_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_TIPO_DOCUMENTO AND dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_CAR_CLASE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_CLASE_ID AND "
            var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO"
            var_cadena = var_cadena + " WHERE (dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_AGE_AGENTE_ID = '" + txt_agente + "') and (dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.vcha_rut_ruta_id = '" + txt_ruta + "')  AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA BETWEEN '" + txt_fecha_inicio + "' AND '" + CStr(var_fecha_fin + 1) + "')"
            var_cadena = var_cadena + " GROUP BY dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_ART_CODIGO_EXTERNO, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_AGE_AGENTE_ID,"
            var_cadena = var_cadena + " dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_CLI_CLAVE_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.AÑO, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.MES, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.DIA, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_CLI_CLAVE_ANTERIOR_ID, dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_DEVOLUCIONES.VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  var_fecha = Format(CStr(rs!mes) + "/" + CStr(rs!dia) + "/" + CStr(rs!año), "mm/dd/yy")
                  var_cadena = "insert into tem_" + var_Archivo + ".dbf (cveestilo, cvecliente, cvetienda, fecha, especial, cantpedi1, cantpedi2, cantpedi3, cantpedi4, cantpedi5, cantpedi6, cantsurt1, cantsurt2, cantsurt3, cantsurt4, cantsurt5, cantsurt6, importe, dcto, finiperiod, ffinperiod) values "
                  var_cadena = var_cadena + " ('" + IIf(IsNull(rs!VCHA_aRT_CODIGO_EXTERNO), "", rs!VCHA_aRT_CODIGO_EXTERNO) + "', '" + IIf(IsNull(rs!vcha_cli_clave_anterior_id), "", rs!vcha_cli_clave_anterior_id) + "', '" + IIf(IsNull(rs!vcha_esb_establecimiento_anterior_id), "", rs!vcha_esb_establecimiento_anterior_id) + "', CTOD('" + CStr(rs!mes) + "/" + CStr(rs!dia) + "/" + CStr(rs!año) + "'), .F.,0, "
                  var_cadena = var_cadena + " 0,0,0,0,0," + CStr(IIf(IsNull(rs!CANTIDAD_DEVUELTA), 0, rs!CANTIDAD_DEVUELTA)) + ", 0,0,0,0,0,0,'D',ctod('" + Format(CDate(Me.txt_fecha_inicio), "mm/dd/yy") + "'), ctod('" + Format(CDate(Me.txt_fecha_fin), "mm/dd/yy") + "'))"
                  rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
               Wend
            End If
            rs.Close
            var_cadena = "SELECT VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_AGE_AGENTE_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, FECHA, IMPORTE_NETO , SUBIMPORTE, año, mes, dia, VCHA_CLI_CLAVE_ANTERIOR_ID, VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID From dbo.VW_AUXILIAR_PEDIDO_SUGERIDO_IMPORTES "
            var_cadena = var_cadena + " WHERE (VCHA_AGE_AGENTE_ID = '" + txt_agente + "') AND (FECHA >= '" + txt_fecha_inicio + "' AND FECHA <= '" + CStr(var_fecha_fin) + "') AND (VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') and vcha_rut_ruta_id = '" + txt_ruta + "'"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  var_cadena = "insert into tem_" + var_Archivo + ".dbf (cveestilo, cvecliente, cvetienda, fecha, especial, cantpedi1, cantpedi2, cantpedi3, cantpedi4, cantpedi5, cantpedi6, cantsurt1, cantsurt2, cantsurt3, cantsurt4, cantsurt5, cantsurt6, importe, dcto,finiperiod, ffinperiod) values "
                  var_cadena = var_cadena + " ('VTANETA', '" + IIf(IsNull(rs!vcha_cli_clave_anterior_id), "", rs!vcha_cli_clave_anterior_id) + "', '" + IIf(IsNull(rs!vcha_esb_establecimiento_anterior_id), "", rs!vcha_esb_establecimiento_anterior_id) + "', CTOD('" + CStr(rs!mes) + "/" + CStr(rs!dia) + "/" + CStr(rs!año) + "'), .F.,0, "
                  var_cadena = var_cadena + " 0,0,0,0,0,0,0,0,0,0,0," + CStr(IIf(IsNull(rs!SUBIMPORTE), 0, rs!SUBIMPORTE)) + ",'V',ctod('" + Format(CDate(Me.txt_fecha_inicio), "mm/dd/yy") + "'), ctod('" + Format(CDate(Me.txt_fecha_fin), "mm/dd/yy") + "'))"
                  rsaux.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
               Wend
            End If
            rs.Close
            rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_correo_electronico = ""
            If Not rs.EOF Then
               var_correo_electronico = IIf(IsNull(rs!VCHA_AGE_EMAIL), "", rs!VCHA_AGE_EMAIL)
            End If
            rs.Close
            
            
            
            If Dir(var_ruta & "\" + var_Archivo + ".dbf") <> "" Then
               rs.Open "DELETE FROM " + var_Archivo, var_tabla, adOpenDynamic, adLockOptimistic
               Kill var_ruta & "\" + var_Archivo + ".dbf"
            End If
            If Dir(var_ruta & "\CLIENTES.dbf") <> "" Then
               rs.Open "DELETE FROM clientes", var_tabla, adOpenDynamic, adLockOptimistic
               Kill var_ruta & "\CLIENTES.dbf"
            End If
            If Dir(var_ruta & "\DETATIEN.dbf") <> "" Then
               rs.Open "DELETE FROM DETATIEN", var_tabla, adOpenDynamic, adLockOptimistic
               Kill var_ruta & "\DETATIEN.dbf"
            End If
            If Dir(var_ruta & "\TIENDAS.dbf") <> "" Then
               rs.Open "DELETE FROM TIENDAS", var_tabla, adOpenDynamic, adLockOptimistic
               Kill var_ruta & "\TIENDAS.dbf"
            End If
            If Dir(var_ruta & "\TITULAR.dbf") <> "" Then
               rs.Open "DELETE FROM TITULAR", var_tabla, adOpenDynamic, adLockOptimistic
               Kill var_ruta & "\TITULAR.dbf"
            End If
           
            var_copia = CopyFile(var_ruta & "tem_clientes.dbf", var_ruta & "clientes.dbf", 1)
            var_copia = CopyFile(var_ruta & "tem_detatien.dbf", var_ruta & "detatien.dbf", 1)
            var_copia = CopyFile(var_ruta & "tem_tiendas.dbf", var_ruta & "tiendas.dbf", 1)
            var_copia = CopyFile(var_ruta & "tem_titular.dbf", var_ruta & "titular.dbf", 1)
            var_copia = CopyFile(var_ruta & "tem_" + var_Archivo + ".dbf", var_ruta + var_Archivo + ".dbf", 1)
            
            
            var_tabla.Close
            
            If Trim(var_correo_electronico) <> "" Then
               If MAPISession1.SessionID = 0 Then
                  MAPISession1.SignOn
               End If
               MAPIMessages1.SessionID = MAPISession1.SessionID
               MAPIMessages1.Compose
               MAPIMessages1.RecipDisplayName = var_correo_electronico
               MAPIMessages1.RecipAddress = var_correo_electronico
               MAPIMessages1.AddressResolveUI = True
               MAPIMessages1.ResolveName
               MAPIMessages1.MsgIndex = -1
               'MAPIMessages1.AttachmentIndex = 4
               'MAPIMessages1.AttachmentCount = 4
               
               
               
               MAPIMessages1.MsgSubject = "Archivos para pedido sugerido (" + Trim(Me.txt_nombre_ruta) + ")"
               MAPIMessages1.MsgNoteText = "Se adjunta archivos para pedido sugerido y para actualizacion de clientes en el sistema de relación de cobranza."
               'For var_i = 0 To 4
               '   If var_i = 0 Then
                     MAPIMessages1.AttachmentIndex = 0 ' numero del anexo, 0,1,2,3....
                     MAPIMessages1.AttachmentType = 0
                     MAPIMessages1.AttachmentName = "clientes.dbf"
                     MAPIMessages1.AttachmentPathName = var_ruta + "clientes.dbf"
               '   End If
               '   If var_i = 1 Then
                     MAPIMessages1.AttachmentIndex = 1 ' numero del anexo, 0,1,2,3....
                     MAPIMessages1.AttachmentType = 0
                     MAPIMessages1.AttachmentName = "tiendas.dbf"
                     MAPIMessages1.AttachmentPathName = var_ruta + "tiendas.dbf"
               '   End If
               '   If var_i = 2 Then
                     MAPIMessages1.AttachmentIndex = 2 ' numero del anexo, 0,1,2,3....
                     MAPIMessages1.AttachmentType = 0
                     MAPIMessages1.AttachmentName = "titular.dbf"
                     MAPIMessages1.AttachmentPathName = var_ruta + "titular.dbf"
               '   End If
               '   If var_i = 3 Then
                     MAPIMessages1.AttachmentIndex = 3 ' numero del anexo, 0,1,2,3....
                     MAPIMessages1.AttachmentType = 0
                     MAPIMessages1.AttachmentName = "detatien.dbf"
                     MAPIMessages1.AttachmentPathName = var_ruta + "detatien.dbf"
               '   End If
               '   If var_i = 4 Then
                     MAPIMessages1.AttachmentIndex = 4 ' numero del anexo, 0,1,2,3....
                     MAPIMessages1.AttachmentType = 0
                     MAPIMessages1.AttachmentName = var_Archivo + ".dbf"
                     MAPIMessages1.AttachmentPathName = var_ruta + var_Archivo + ".dbf"
               '   End If
               'Next var_i
               MAPIMessages1.Send True
               If MAPISession1.SessionID > 0 Then
                  MAPISession1.SignOff
               End If
            Else
               MsgBox "Debe de indicar una cuenta de correo al agente en el catálogo de clientes", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La fecha final es incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La fecha de inicio es incorrecta", vbOKOnly, "ATENCION"
      End If
      Else
        MsgBox "No se a seleccionado una ruta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_mes_1_Click()
   mon_mes1.Value = Date
   mon_mes1.Visible = True
   mon_mes1.SetFocus
End Sub

Private Sub cmd_mes_2_Click()
   mon_mes2.Value = Date
   mon_mes2.Visible = True
   mon_mes2.SetFocus
End Sub

Private Sub cmd_nuevo_Click()
  Me.txt_agente = ""
  Me.txt_nombre_agente = ""
  Me.txt_ruta = ""
  Me.txt_nombre_ruta = ""
  txt_agente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2000
   Left = 3200
   frm_lista.Visible = False
   txt_fecha_inicio = Date
   txt_fecha_fin = Date
   mon_mes1.Visible = False
   mon_mes2.Visible = False
   Set var_tabla = CreateObject("ADODB.connection")
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select VCHA_PRI_RUTA_PEDIDO_SUGERIDO from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   var_ruta = IIf(IsNull(rs(0).Value), "", rs(0).Value)
   rs.Close
   If Trim(var_ruta) <> "" Then
      If Not UCase(VAR_MAQUINA) = "JFSERNA" Then
         var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
       Else
         var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
       End If
   Else
      MsgBox "No se a indicado una ruta para los archivos a enviar", vbOKOnly, "ATENCION"
      txt_agente.Enabled = False
      cmb_agentes.Enabled = False
      txt_fecha_inicio.Enabled = False
      txt_fecha_fin.Enabled = False
      cmd_mes_1.Enabled = False
      cmd_mes_2.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_informacion_pedido_sugerido)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            Me.txt_agente = lv_lista.selectedItem
            Me.txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         End If
         Me.txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            Me.txt_ruta = lv_lista.selectedItem
            Me.txt_nombre_ruta = lv_lista.selectedItem.SubItems(1)
         End If
         Me.txt_ruta.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub mon_mes1_DateDblClick(ByVal DateDblClicked As Date)
   txt_fecha_inicio = mon_mes1.Value
   mon_mes1.Visible = False
End Sub

Private Sub mon_mes1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fecha_inicio = mon_mes1.Value
      mon_mes1.Visible = False
   End If
   If KeyAscii = 27 Then
      mon_mes1.Visible = False
   End If
End Sub

Private Sub mon_mes1_LostFocus()
   mon_mes1.Visible = False
End Sub

Private Sub mon_mes2_DateDblClick(ByVal DateDblClicked As Date)
   txt_fecha_fin = mon_mes2.Value
   mon_mes2.Visible = False
   txt_fecha_fin.SetFocus
End Sub

Private Sub mon_mes2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fecha_fin = mon_mes2.Value
      mon_mes2.Visible = False
   End If
   If KeyAscii = 27 Then
      mon_mes2.Visible = False
   End If
End Sub

Private Sub mon_mes2_LostFocus()
   mon_mes2.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = rs!VCHA_AGE_NOMBRE
         Me.txt_ruta = ""
         Me.txt_nombre_ruta = ""
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
         Me.txt_ruta = ""
         Me.txt_nombre_ruta = ""
      End If
      rs.Close
   Else
      txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_rut_ruta_id, vcha_rut_nombre from TB_RUTAS where  vcha_age_agente_id = '" + txt_agente + "' order by vcha_RUT_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_rut_ruta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "RUTAS"
      var_tipo_lista = 2
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

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_LostFocus()
   If Trim(txt_ruta) <> "" Then
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
      Else
         MsgBox "Clave de ruta incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_ruta = ""
         txt_ruta = ""
      End If
      rs.Close
   Else
      txt_nombre_ruta = ""
   End If
End Sub
