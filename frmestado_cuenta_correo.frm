VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmestado_cuenta_correo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio masivo de estado de cuentas"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   930
      Top             =   0
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
      Left            =   1485
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame frmlista_correo 
      Height          =   2400
      Left            =   60
      TabIndex        =   25
      Top             =   465
      Width           =   7425
      Begin MSComctlLib.ListView lv_lista_correo 
         Height          =   1830
         Left            =   60
         TabIndex        =   26
         Top             =   465
         Width           =   7275
         _ExtentX        =   12832
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
            Text            =   "Cliente"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Correo"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_correo 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   135
         Width           =   7350
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1350
      TabIndex        =   21
      Top             =   420
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7050
      Picture         =   "frmestado_cuenta_correo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmestado_cuenta_correo.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Enviar estado de cuenta por correo"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   60
      TabIndex        =   18
      Top             =   300
      Width           =   7440
   End
   Begin VB.Frame Frame2 
      Caption         =   " Periodo "
      Height          =   690
      Left            =   120
      TabIndex        =   15
      Top             =   2775
      Width           =   7230
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   2100
         TabIndex        =   11
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1620
         TabIndex        =   17
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3765
         TabIndex        =   16
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cliente "
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   465
      Width           =   7230
      Begin VB.TextBox txt_agente 
         Height          =   345
         Left            =   1035
         TabIndex        =   2
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   345
         Left            =   2310
         TabIndex        =   3
         Top             =   240
         Width           =   4785
      End
      Begin VB.TextBox txt_correo 
         Height          =   345
         Left            =   1035
         TabIndex        =   10
         Top             =   1770
         Width           =   6060
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   345
         Left            =   2310
         TabIndex        =   9
         Top             =   1380
         Width           =   4785
      End
      Begin VB.TextBox txt_cliente 
         Height          =   345
         Left            =   1035
         TabIndex        =   8
         Top             =   1380
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   345
         Left            =   2310
         TabIndex        =   7
         Top             =   990
         Width           =   4785
      End
      Begin VB.TextBox txt_titular 
         Height          =   345
         Left            =   1035
         TabIndex        =   6
         Top             =   1005
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_grupo 
         Height          =   345
         Left            =   2310
         TabIndex        =   5
         Top             =   615
         Width           =   4785
      End
      Begin VB.TextBox txt_grupo 
         Height          =   345
         Left            =   1035
         TabIndex        =   4
         Top             =   615
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Correo:"
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   1830
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1455
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   1065
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   690
         Width           =   480
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages2 
      Left            =   2175
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession2 
      Left            =   2730
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmestado_cuenta_correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_lista As Integer

Private Sub cmd_correo_Click()
   'On Error GoTo salir:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If rsaux9.State = 1 Then
      rsaux9.Close
   End If
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If Me.txt_agente <> "" Then
            cnn.CommandTimeout = 360
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ESTADO_CUENTA_CORREO"
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 1
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_ESTADO_CUENTA_CORREO (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
          
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(Me.txt_fin))
            var_mes = CStr(Month(Me.txt_fin))
            var_año = CStr(Year(Me.txt_fin))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        
            rs.Open "EXEC SP_ESTADO_CUENTA_CORREO_agente " + CStr(var_consecutivo) + "," + var_fecha_inicio + ", " + var_fecha_fin + ",'" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.State = 1 Then
               rs.Close
            End If
            rsaux10.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre from TB_TEMP_ESTADO_CUENTA_CORREO  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " And vcha_tit_titular_id Is Not Null", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  rsaux11.Open "select distinct vcha_cli_clave_id from TB_TEMP_ESTADO_CUENTA_CORREO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_tit_titular_id = '" + rsaux10!vcha_tit_titular_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  VAR_PRIMERO_POSITIVO = 1
                  VAR_PRIMERO_NEGATIVO = 1
                  var_index_1 = 0
                  var_index_2 = 0
                  While Not rsaux11.EOF
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "SELECT  vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial, sum(floa_tem_importe_cargo) AS IMPORTE_cARGOS, SUM(floa_tem_importe_abono) AS IMPORTE_ABONOS  FROM TB_TEMP_ESTADO_CUENTA_CORREO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_tit_titular_id = '" + IIf(IsNull(rsaux10!vcha_tit_titular_id), "", rsaux10!vcha_tit_titular_id) + "' AND VCHA_CLI_CLAVE_ID = '" + IIf(IsNull(rsaux11!VCHA_CLI_CLAVE_ID), "", rsaux11!VCHA_CLI_CLAVE_ID) + "' group by vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial", cnn, adOpenDynamic, adLockOptimistic
                        If VAR_PRIMERO_POSITIVO = 1 Then
                           If Not rs.EOF Then
                              rsaux.Open "select top 1 isnull(vcha_cli_email,0) as correo from tb_clientes where vcha_tit_titular_id = '" + rsaux10!vcha_tit_titular_id + "' and isnull(vcha_cli_email,'') <> ''", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 Me.txt_correo = IIf(IsNull(rsaux!correo), "", rsaux!correo)
                              Else
                                 Me.txt_correo = ""
                              End If
                              rsaux.Close
                           Else
                              Me.txt_correo = ""
                           End If
                        End If
                        If Me.txt_correo <> "" Then
                           var_cadena_clientes = ""
                           var_cadena_clientes_negativo = ""
                           While Not rs.EOF
                                 var_saldo = IIf(IsNull(rs!floa_tem_saldo_inicial), 0, rs!floa_tem_saldo_inicial)
                                 var_diferencia = IIf(IsNull(rs!IMPORTE_CARGOS), 0, rs!IMPORTE_CARGOS) - IIf(IsNull(rs!importe_abonos), 0, rs!importe_abonos)
                                 var_saldo = CDbl(var_saldo + var_diferencia)
                                 If CDbl(var_saldo) >= -10 Then
                                    If var_cadena_clientes = "" Then
                                       var_cadena_clientes = "{VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rsaux11!VCHA_CLI_CLAVE_ID + "'"
                                    Else
                                       var_cadena_clientes = var_cadena_clientes + " or {VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rsaux11!VCHA_CLI_CLAVE_ID + "'"
                                    End If
                                 Else
                                    If var_cadena_clientes_negativo = "" Then
                                       var_cadena_clientes_negativo = "{VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rsaux11!VCHA_CLI_CLAVE_ID + "'"
                                    Else
                                       var_cadena_clientes_negativo = var_cadena_clientes_negativo + " or {VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rsaux11!VCHA_CLI_CLAVE_ID + "'"
                                    End If
                                 End If
                                 rs.MoveNext
                           Wend
                           var_si_informacion = 1
                           If rs.RecordCount > 0 Then
                              rs.MoveFirst
                              var_si_informacion = 1
                           Else
                              var_si_informacion = 0
                           End If
                           'var_correo = "fserna@vianney.com.mx"
                           var_correo = Me.txt_correo
                           If var_cadena_clientes <> "" Then
                              Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_CORREO.rpt")
                              reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_CORREO_REPORTE.INTE_TEM_CONSECUTIVO}=" + CStr(var_consecutivo) + " and (" + var_cadena_clientes + ")"
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              reporte.ExportOptions.FormatType = crEFTExcel80
                              reporte.ExportOptions.DestinationType = crEDTDiskFile
                              var_nombre_archivo = "_CLIENTE_" + rs!VCHA_CLI_CLAVE_ID + "_" + rs!VCHA_CLI_NOMBRE
                              var_subject = "Estado de cuenta del titular " + rsaux10!VCHA_TIT_NOMBRE
                              var_nota = "Se adjunta archivo con el estado de cuenta del titular : " + rsaux10!VCHA_TIT_NOMBRE
                              archivo = "c:\reportessid\ESTADO_CUENTA_" + var_nombre_archivo + ".xls"
                              reporte.ExportOptions.DiskFileName = archivo
                              reporte.Export False
                              Set reporte = Nothing
                              var_cadena_correo = var_correo
                              If VAR_PRIMERO_POSITIVO = 1 Then
                                 If MAPISession1.SessionID = 0 Then
                                    MAPISession1.SignOn
                                 End If
                                 MAPIMessages1.SessionID = MAPISession1.SessionID
                                 MAPIMessages1.Compose
                                 MAPIMessages1.RecipDisplayName = var_cadena_correo
                                 MAPIMessages1.RecipAddress = var_cadena_correo
                                 MAPIMessages1.AddressResolveUI = True
                                 MAPIMessages1.ResolveName
                                 MAPIMessages1.MsgSubject = var_subject
                                 var_nota = "Estimado cliente: "
                                 MAPIMessages1.MsgNoteText = var_nota
                                 var_nota = var_nota + Chr(13) + " Favor de validar su información y confirmar por esta misma vía cualquier duda o comentario sobre el mismo. " + Chr(13)
                                 MAPIMessages1.MsgNoteText = var_nota
                                 var_nota = var_nota + Chr(13) + " De antemano le agradecemos su atención. " + Chr(13) + " Departamento de crédito y cobranza" + Chr(13) + " Grupo vianney."


                                 MAPIMessages1.MsgNoteText = var_nota
                                 VAR_PRIMERO_POSITIVO = 2
                              End If
                              MAPIMessages1.AttachmentIndex = var_index_1
                              MAPIMessages1.AttachmentPathName = archivo
                              var_index_1 = var_index_1 + 1

                           End If
                           If var_cadena_clientes_negativo <> "" Then
                              var_correo = "vcovarrubias@vianney.com.mx"
                              Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_CORREO.rpt")
                              reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_CORREO_REPORTE.INTE_TEM_CONSECUTIVO}=" + CStr(var_consecutivo) + " and (" + var_cadena_clientes_negativo + ")"
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              reporte.ExportOptions.FormatType = crEFTExcel80
                              reporte.ExportOptions.DestinationType = crEDTDiskFile
                              
                              archivo = "c:\reportessid\ESTADO_CUENTA_" + var_nombre_archivo + ".xls"
                              reporte.ExportOptions.DiskFileName = archivo
                              reporte.Export False
                              Set reporte = Nothing
                              var_cadena_correo = var_correo
                              'rsaux.Open "IN_PKG_CORREOS_2.SP_ENVIAR_EMAIL('asifuentes@vianney.com.mx','" + var_correo + "', 'asifuentes@vianney.com.mx', '', 'Estado de cuenta del cliente " + Me.txt_nombre_cliente + "','Se adjunta archivo con el estado de cuenta del grupo : " + Me.txt_nombre_grupo + "','ESTADO_CUENTA_" + var_nombre_archivo + ".xls')", cnnoracle, adOpenDynamic, adLockOptimistic
                         
                              If VAR_PRIMERO_NEGATIVO = 1 Then
                                 If MAPISession2.SessionID = 0 Then
                                    MAPISession2.SignOn
                                 End If
                                 MAPIMessages2.SessionID = MAPISession1.SessionID
                                 MAPIMessages2.Compose
                                 MAPIMessages2.RecipDisplayName = var_cadena_correo
                                 MAPIMessages2.RecipAddress = var_cadena_correo
                                 MAPIMessages2.AddressResolveUI = True
                                 MAPIMessages2.ResolveName
                                 MAPIMessages2.MsgSubject = var_subject
                                 MAPIMessages2.MsgNoteText = var_nota
                                 VAR_PRIMERO_NEGATIVO = 2
                              End If
                              MAPIMessages1.AttachmentIndex = var_index_2
                              MAPIMessages2.AttachmentPathName = archivo
                              var_index_2 = var_index_2 + 1
                           End If
                        End If
                        rsaux11.MoveNext
                  Wend
                  rsaux11.Close
                  If VAR_PRIMERO_NEGATIVO = 2 Then
                     MAPIMessages2.Send False
                     If MAPISession2.SessionID > 0 Then
                        MAPISession2.SignOff
                     End If
                  End If
                  
                  If VAR_PRIMERO_POSITIVO = 2 Then
                     MAPIMessages1.Send False
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
                  rs.Close
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
            MsgBox "Se a terminado el proceso de envio de correos", vbOKOnly, "ATENCION"
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "DELETE FROM TB_TEMP_ESTADO_CUENTA_CORREO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            If Me.txt_grupo <> "" Then
               cnn.CommandTimeout = 360
               cnn.BeginTrans
               rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ESTADO_CUENTA_CORREO"
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               Else
                  var_consecutivo = 1
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "INSERT INTO TB_TEMP_ESTADO_CUENTA_CORREO (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
          
         
               var_dia = CStr(Day(CDate(txt_inicio)))
               var_mes = CStr(Month(CDate(txt_inicio)))
               var_año = CStr(Year(CDate(txt_inicio)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
               var_dia = CStr(Day(Me.txt_fin))
               var_mes = CStr(Month(Me.txt_fin))
               var_año = CStr(Year(Me.txt_fin))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                         
               rs.Open "EXEC SP_ESTADO_CUENTA_CORREO " + CStr(var_consecutivo) + "," + var_fecha_inicio + ", " + var_fecha_fin + ",'" + Me.txt_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Me.txt_cliente <> "" Then
                  rs.Open "SELECT  vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial, sum(floa_tem_importe_cargo) AS IMPORTE_cARGOS, SUM(floa_tem_importe_abono) AS IMPORTE_ABONOS FROM TB_TEMP_ESTADO_CUENTA_CORREO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_cli_clave_id = '" + Me.txt_cliente + "' group by vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial", cnn, adOpenDynamic, adLockOptimistic
               Else
                  If Me.txt_titular <> "" Then
                     rs.Open "SELECT  vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial, sum(floa_tem_importe_cargo) AS IMPORTE_cARGOS, SUM(floa_tem_importe_abono) AS IMPORTE_ABONOS  FROM TB_TEMP_ESTADO_CUENTA_CORREO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_tit_titular_id = '" + Me.txt_titular + "' group by vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rs.Open "SELECT  vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial, sum(floa_tem_importe_cargo) AS IMPORTE_cARGOS, SUM(floa_tem_importe_abono) AS IMPORTE_ABONOS  FROM TB_TEMP_ESTADO_CUENTA_CORREO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_gac_grupo_actual_id = '" + Me.txt_grupo + "' group by vcha_cli_clave_id, vcha_cli_nombre, floa_tem_saldo_inicial", cnn, adOpenDynamic, adLockOptimistic
                  End If
               End If
            
                        
            
               
               If Me.txt_correo <> "" Then
                  var_cadena_clientes = ""
                  var_cadena_clientes_negativo = ""
                  While Not rs.EOF
                        rsaux1.Open "select * from tb_clientes where vcha_Cli_clave_id = '" + rs!VCHA_CLI_CLAVE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                        'var_correo = CStr(IIf(IsNull(rsaux1!VCHA_cLI_EMAIL), "", rsaux1!VCHA_cLI_EMAIL))
                        rsaux1.Close
                        var_saldo = IIf(IsNull(rs!floa_tem_saldo_inicial), 0, rs!floa_tem_saldo_inicial)
                        var_diferencia = IIf(IsNull(rs!IMPORTE_CARGOS), 0, rs!IMPORTE_CARGOS) - IIf(IsNull(rs!importe_abonos), 0, rs!importe_abonos)
                        var_saldo = CDbl(var_saldo + var_diferencia)
                        If CDbl(var_saldo) >= -10 Then
                           If var_cadena_clientes = "" Then
                              var_cadena_clientes = "{VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rs!VCHA_CLI_CLAVE_ID + "'"
                           Else
                              var_cadena_clientes = var_cadena_clientes + " or {VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rs!VCHA_CLI_CLAVE_ID + "'"
                           End If
                        Else
                           If var_cadena_clientes_negativo = "" Then
                              var_cadena_clientes_negativo = "{VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rs!VCHA_CLI_CLAVE_ID + "'"
                           Else
                              var_cadena_clientes_negativo = var_cadena_clientes_negativo + " or {VW_ESTADO_CUENTA_CORREO_REPORTE.VCHA_CLI_CLAVE_ID} = '" + rs!VCHA_CLI_CLAVE_ID + "'"
                           End If
                        End If
                        rs.MoveNext
                  Wend
                  var_si_informacion = 1
                  If rs.RecordCount > 0 Then
                     rs.MoveFirst
                     var_si_informacion = 1
                  Else
                     var_si_informacion = 0
                  End If
                  'var_correo = "fserna@vianney.com.mx"
                  var_correo = Me.txt_correo
                  If var_cadena_clientes <> "" Then
                     Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_CORREO.rpt")
                     reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_CORREO_REPORTE.INTE_TEM_CONSECUTIVO}=" + CStr(var_consecutivo) + " and (" + var_cadena_clientes + ")"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     If Trim(Me.txt_grupo) <> "" Then
                        If Trim(Me.txt_titular) <> "" Then
                           If Trim(Me.txt_cliente) <> "" Then
                              var_nombre_archivo = "_CLIENTE_" + Me.txt_cliente
                              var_subject = "Estado de cuenta del cliente " + Me.txt_nombre_cliente
                              var_nota = "Se adjunta archivo con el estado de cuenta del cliente : " + Me.txt_nombre_cliente
                           Else
                              var_nombre_archivo = "_TITULAR_" + Me.txt_titular
                              var_subject = "Estado de cuenta del titular " + Me.txt_nombre_titular
                              var_nota = "Se adjunta archivo con el estado de cuenta del titular : " + Me.txt_nombre_titular
                           End If
                        Else
                           var_nombre_archivo = "_GRUPO_" + Me.txt_grupo
                           var_subject = "Estado de cuenta del grupo " + Me.txt_nombre_grupo
                           var_nota = "Se adjunta archivo con el estado de cuenta del grupo : " + Me.txt_nombre_grupo
                        End If
                     End If
                   
                     'archivo = "\\Administra\ERPOracle\OC\ESTADO_CUENTA_" + var_nombre_archivo + ".xls"
                     archivo = "c:\reportessid\ESTADO_CUENTA_" + var_nombre_archivo + ".xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     'rsaux.Open "IN_PKG_CORREOS_2.SP_ENVIAR_EMAIL('asifuentes@vianney.com.mx','" + var_correo + "', 'asifuentes@vianney.com.mx', '', 'Estado de cuenta del cliente " + Me.txt_nombre_cliente + "','Se adjunta archivo con el estado de cuenta del grupo : " + Me.txt_nombre_grupo + "','ESTADO_CUENTA_" + var_nombre_archivo + ".xls')", cnnoracle, adOpenDynamic, adLockOptimistic
                     var_cadena_correo = var_correo
                    
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = var_cadena_correo
                     MAPIMessages1.RecipAddress = var_cadena_correo
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = var_subject
                     MAPIMessages1.MsgNoteText = var_nota
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.Send False
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
                  If var_cadena_clientes_negativo <> "" Then
                     var_correo = "vcovarrubias@vianney.com.mx"
                     Set reporte = appl.OpenReport(App.Path + "\REP_ESTADO_CUENTA_CORREO.rpt")
                     reporte.RecordSelectionFormula = "{VW_ESTADO_CUENTA_CORREO_REPORTE.INTE_TEM_CONSECUTIVO}=" + CStr(var_consecutivo) + " and (" + var_cadena_clientes_negativo + ")"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     If Trim(Me.txt_grupo) <> "" Then
                        If Trim(Me.txt_titular) <> "" Then
                           If Trim(Me.txt_cliente) <> "" Then
                              var_nombre_archivo = "_CLIENTE_" + Me.txt_cliente
                              var_subject = "Estado de cuenta del cliente " + Me.txt_nombre_cliente
                              var_nota = "Se adjunta archivo con el estado de cuenta del cliente : " + Me.txt_nombre_cliente
                           Else
                              var_nombre_archivo = "_TITULAR_" + Me.txt_titular
                              var_subject = "Estado de cuenta del titular " + Me.txt_nombre_titular
                              var_nota = "Se adjunta archivo con el estado de cuenta del titular : " + Me.txt_nombre_titular
                           End If
                        Else
                           var_nombre_archivo = "_GRUPO_" + Me.txt_grupo
                           var_subject = "Estado de cuenta del grupo " + Me.txt_nombre_grupo
                           var_nota = "Se adjunta archivo con el estado de cuenta del grupo : " + Me.txt_nombre_grupo
                        End If
                     End If
                        
                     archivo = "c:\reportessid\ESTADO_CUENTA_" + var_nombre_archivo + ".xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     var_cadena_correo = var_correo
                     'rsaux.Open "IN_PKG_CORREOS_2.SP_ENVIAR_EMAIL('asifuentes@vianney.com.mx','" + var_correo + "', 'asifuentes@vianney.com.mx', '', 'Estado de cuenta del cliente " + Me.txt_nombre_cliente + "','Se adjunta archivo con el estado de cuenta del grupo : " + Me.txt_nombre_grupo + "','ESTADO_CUENTA_" + var_nombre_archivo + ".xls')", cnnoracle, adOpenDynamic, adLockOptimistic
                     
                         
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = var_cadena_correo
                     MAPIMessages1.RecipAddress = var_cadena_correo
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = var_subject
                     MAPIMessages1.MsgNoteText = var_nota
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.Send False
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
               End If
               MsgBox "Se a terminado el proceso de envio de correos", vbOKOnly, "ATENCION"
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "DELETE FROM TB_TEMP_ESTADO_CUENTA_CORREO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "No se a seleccionado un grupo", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Fecha fin incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecto", vbOKOnly, "ATENCION"
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If rsaux9.State = 1 Then
      rsaux9.Close
   End If
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
Exit Sub
salir:
   MsgBox "A surgido un error al generar el correo", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If rsaux9.State = 1 Then
      rsaux9.Close
   End If
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2000
   Left = 2000
   Me.txt_fin = Date
   Me.txt_inicio = Date
   frm_lista.Visible = False
   Me.frmlista_correo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_correo_KeyPress(KeyAscii As Integer)
   If Me.lv_lista_correo.ListItems.Count > 0 Then
      If KeyAscii = 13 Then
         Me.txt_correo = Me.lv_lista_correo.selectedItem.SubItems(1)
         Me.txt_correo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_correo.SetFocus
   End If
End Sub

Private Sub lv_lista_correo_LostFocus()
   Me.frmlista_correo.Visible = False
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If Me.lv_lista.ListItems.Count > 0 Then
            Me.txt_grupo = lv_lista.selectedItem
            Me.txt_nombre_grupo = lv_lista.selectedItem.SubItems(1)
            Me.txt_grupo.SetFocus
         End If
      End If
      If var_tipo_lista = 2 Then
         If Me.lv_lista.ListItems.Count > 0 Then
            Me.txt_titular = lv_lista.selectedItem
            Me.txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
            Me.txt_titular.SetFocus
         End If
      End If
      If var_tipo_lista = 3 Then
         If Me.lv_lista.ListItems.Count > 0 Then
            Me.txt_cliente = lv_lista.selectedItem
            Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            Me.txt_cliente.SetFocus
         End If
      End If
      If var_tipo_lista = 4 Then
         If Me.lv_lista.ListItems.Count > 0 Then
            Me.txt_agente = lv_lista.selectedItem
            Me.txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
            Me.txt_agente.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         Me.txt_grupo.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_titular.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_cliente.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub Text1_Change()

End Sub


Private Sub txt_agente_Change()
   Me.txt_nombre_agente = ""
   Me.txt_grupo = ""
   Me.txt_nombre_grupo = ""
   Me.txt_titular = ""
   Me.txt_nombre_titular = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_correo = ""
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_AGE_AGENTE_ID IS NOT NULL and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 4
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre_agente.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      End If
   End If
End Sub

Private Sub txt_agente_LostFocus()
   If Me.txt_agente <> "" Then
      rs.Open "select * from tb_agentes where vcha_Age_agente_id = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "El agente no existe", vbOKOnly, "ATENCION"
         Me.txt_nombre_agente = ""
      End If
      rs.Close
   Else
       Me.txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_cliente_Change()
   Me.txt_nombre_cliente = ""
   Me.txt_correo = ""
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_grupo <> "" Then
         lv_lista.ListItems.Clear
         rs.Open "select distinct VCHA_CLI_CLAVE_ID, vcha_CLI_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "' AND VCHA_GAC_GRUPO_ACTUAL_ID = '" + Me.txt_grupo + "'   and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "CLIENTES"
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 3900.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4130.71
         End If
         var_tipo_lista = 3
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre_cliente.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_cliente_LostFocus()
   If Me.txt_grupo <> "" Then
      If Me.txt_titular <> "" Then
         If Me.txt_cliente <> "" Then
            rs.Open "select * from vw_clientes where vcha_gac_grupo_Actual_id = '" + Me.txt_grupo + "' and vcha_tit_titular_id = '" + Me.txt_titular + "' and vcha_Cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            Else
               MsgBox "Clave de cliente no existe o no pertenece al grupo seleccionado", vbOKOnly, "ATENCION"
               Me.txt_nombre_cliente = ""
            End If
            rs.Close
         Else
            Me.txt_nombre_cliente = ""
         End If
      Else
         MsgBox "No se a seleccionado un titular", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un grupo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_correo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Dim var_n As Integer
      If Me.txt_grupo <> "" Then
         If Me.txt_titular <> "" Then
            If Me.txt_cliente <> "" Then
               lv_lista_correo.ListItems.Clear
               rs.Open "select distinct vcha_cli_nombre, vcha_cli_email from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_GAC_GRUPO_aCTUAL_Id = '" + Me.txt_grupo + "' and isnull(vcha_cli_email,'') <> '' and vcha_tit_titular_id = '" + Me.txt_titular + "' and vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockBatchOptimistic
               While Not rs.EOF
                     Set list_item = lv_lista_correo.ListItems.Add(, , rs!VCHA_CLI_NOMBRE)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     rs.MoveNext
               Wend
               rs.Close
               lbl_lista = "CORREOS"
               var_n = lv_lista_correo.ListItems.Count
               var_tipo_lista = 1
               Me.frmlista_correo.Visible = True
               Me.lv_lista_correo.SetFocus
            Else
               lv_lista_correo.ListItems.Clear
               rs.Open "select distinct vcha_cli_nombre, vcha_cli_email from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_GAC_GRUPO_aCTUAL_Id = '" + Me.txt_grupo + "' and isnull(vcha_cli_email,'') <> '' and vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockBatchOptimistic
               While Not rs.EOF
                     Set list_item = lv_lista_correo.ListItems.Add(, , rs!VCHA_CLI_NOMBRE)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                     rs.MoveNext
               Wend
               rs.Close
               lbl_lista = "CORREOS"
               var_n = lv_lista_correo.ListItems.Count
               var_tipo_lista = 1
               Me.frmlista_correo.Visible = True
               Me.lv_lista_correo.SetFocus
            End If
         Else
            lv_lista_correo.ListItems.Clear
            rs.Open "select distinct vcha_cli_nombre, vcha_cli_email from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_GAC_GRUPO_aCTUAL_Id = '" + Me.txt_grupo + "' and isnull(vcha_cli_email,'') <> ''", cnn, adOpenDynamic, adLockBatchOptimistic
            While Not rs.EOF
                  Set list_item = lv_lista_correo.ListItems.Add(, , rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email)
                  rs.MoveNext
            Wend
            rs.Close
            lbl_lista = "CORREOS"
            var_n = lv_lista_correo.ListItems.Count
            var_tipo_lista = 1
            Me.frmlista_correo.Visible = True
            Me.lv_lista_correo.SetFocus
         End If
      Else
         MsgBox "No se a seleccionado algún grupo", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_correo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_inicio.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_grupo_Change()
   Me.txt_nombre_grupo = ""
   Me.txt_titular = ""
   Me.txt_nombre_titular = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_correo = ""
End Sub

Private Sub txt_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_gac_grupo_actual_id, vcha_gac_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_GAC_GRUPO_aCTUAL_ID IS NOT NULL and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_grupo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre_grupo.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      End If
   End If
End Sub

Private Sub txt_grupo_LostFocus()
   If Me.txt_grupo <> "" Then
      rs.Open "select * from tb_gruposactuales where vcha_gac_grupo_Actual_id = '" + Me.txt_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_grupo = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
         Me.txt_titular = ""
         Me.txt_nombre_titular = ""
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
      Else
         MsgBox "Clave de grupo no existe", vbOKOnly, "ATENCION"
         Me.txt_nombre_grupo = ""
         Me.txt_titular = ""
         Me.txt_nombre_titular = ""
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_grupo = ""
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_nombre_agente_Change()
   Me.txt_correo = ""
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_AGE_AGENTE_ID IS NOT NULL and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 4
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_grupo.SetFocus
   Else
     If KeyAscii = 27 Then
        Unload Me
     Else
        KeyAscii = 0
     End If
   End If
End Sub

Private Sub txt_nombre_cliente_Change()
   Me.txt_correo = ""
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_grupo <> "" Then
         lv_lista.ListItems.Clear
         rs.Open "select distinct VCHA_CLI_CLAVE_ID, vcha_CLI_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "' AND VCHA_GAC_GRUPO_ACTUAL_ID = '" + Me.txt_grupo + "'   and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLI_CLAVE_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "CLIENTES"
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 3900.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4130.71
         End If
         var_tipo_lista = 3
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_correo.SetFocus
   Else
     If KeyAscii = 27 Then
        Unload Me
     Else
        KeyAscii = 0
     End If
   End If
End Sub

Private Sub txt_nombre_grupo_Change()
   Me.txt_correo = ""
End Sub

Private Sub txt_nombre_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_gac_grupo_actual_id, vcha_gac_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_GAC_GRUPO_aCTUAL_ID IS NOT NULL  and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3900.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4130.71
      End If
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_grupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_titular.SetFocus
   Else
     If KeyAscii = 27 Then
        Unload Me
     Else
        KeyAscii = 0
     End If
   End If
End Sub

Private Sub txt_nombre_titular_Change()
   Me.txt_correo = ""
End Sub

Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_grupo <> "" Then
         lv_lista.ListItems.Clear
         rs.Open "select distinct vcha_tit_titular_id, vcha_TIT_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_TIT_TITULAR_ID IS NOT NULL AND VCHA_GAC_GRUPO_ACTUAL_ID = '" + Me.txt_grupo + "'   and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "TITULARES"
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 3900.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4130.71
         End If
         var_tipo_lista = 2
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cliente.SetFocus
   Else
     If KeyAscii = 27 Then
        Unload Me
     Else
        KeyAscii = 0
     End If
   End If
End Sub

Private Sub txt_titular_Change()
   Me.txt_nombre_titular = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_correo = ""
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_grupo <> "" Then
         lv_lista.ListItems.Clear
         rs.Open "select distinct vcha_tit_titular_id, vcha_TIT_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_TIT_TITULAR_ID IS NOT NULL AND VCHA_GAC_GRUPO_ACTUAL_ID = '" + Me.txt_grupo + "'   and len(isnull(VCHA_cLI_EMAIL,'')) > 0 ", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "TITULARES"
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 3900.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4130.71
         End If
         var_tipo_lista = 2
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
    Me.txt_nombre_titular.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_titular_LostFocus()
   If Me.txt_grupo <> "" Then
      If Me.txt_titular <> "" Then
         rs.Open "select * from vw_clientes where vcha_gac_grupo_Actual_id = '" + Me.txt_grupo + "' and vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            Me.txt_cliente = ""
            Me.txt_nombre_cliente = ""
         Else
            MsgBox "El titular no existe", vbOKOnly, "ATENCION"
            Me.txt_nombre_titular = ""
            Me.txt_titular = ""
            Me.txt_nombre_cliente = ""
            Me.txt_cliente = ""
         End If
         rs.Close
      Else
         Me.txt_nombre_titular = ""
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
      End If
   Else
      MsgBox "No se a seleccionado un grupo", vbOKOnly, "ATENCION"
   End If
End Sub
