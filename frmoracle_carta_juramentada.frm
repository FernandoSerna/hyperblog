VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_carta_juramentada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta juramentada"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3330
      Picture         =   "frmoracle_carta_juramentada.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_carta_juramentada.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_carta_juramentada.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   45
      TabIndex        =   3
      Top             =   405
      Width           =   3720
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1575
         TabIndex        =   4
         Top             =   210
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   5
         Top             =   270
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   1
      Top             =   285
      Width           =   3720
   End
   Begin VB.Frame Frame1 
      Height          =   2550
      Left            =   60
      TabIndex        =   0
      Top             =   1275
      Width           =   3720
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2325
         Left            =   45
         TabIndex        =   2
         Top             =   150
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4101
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Embarque"
            Object.Width           =   5821
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_carta_juramentada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_guardar_Click()
   If Me.lv_lista.ListItems.Count > 0 Then
      var_cadena_facturas = ""
      VAR_TODAS_LAS_FACTURAS = 1
      For var_j = 1 To Me.lv_lista.ListItems.Count
          Me.lv_lista.ListItems.Item(var_j).Selected = True
          Me.txt_embarque = Me.lv_lista.selectedItem
          var_cadena = "SELECT * from xxvia_tb_encabezado_embarques where embarque = ?"
          strconsulta = var_cadena
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_embarque)
               .Parameters.Append parametro
          End With
          Set rsaux12 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
         
          var_posible = 1
          If var_posible = 1 Then
             var_cadena = "SELECT distinct inte_emb_embarque, source_header_number from xxvia_tb_salidas_cajas where inte_emb_embarque = ?"
             strconsulta = var_cadena
             With comandoORA
                  .ActiveConnection = cnnoracle_4
                  .CommandType = adCmdText
                  .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_embarque)
                  .Parameters.Append parametro
             End With
             Set rsaux12 = comandoORA.execute
             Set comandoORA = Nothing
             Set parametro = Nothing
             While Not rsaux12.EOF
                   rsaux1.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   var_cadena = "SELECT b.vcha_tit_titular_id, D.vcha_esb_establecimient_id FROM XXVIA_VW_CLIENTES_PEDIDOS B, OE_ORDER_HEADERS_ALL C, XXVIA_VW_ESTABLECIMIENTOS_PED D, XXVIA_VW_ESTABLECIMIENTOS_PED E Where c.order_number = ? AND c.SOLD_TO_ORG_ID = B.CUST_ACCOUNT_ID AND D.SITE_USE_ID    = C.SHIP_TO_ORG_ID AND E.SITE_USE_ID    = C.INVOICE_TO_ORG_ID"
                   strconsulta = var_cadena
                   With comandoORA
                        .ActiveConnection = cnnoracle_4
                        .CommandType = adCmdText
                        .CommandText = strconsulta
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux12!source_header_number))
                        .Parameters.Append parametro
                   End With
                   Set rsaux9 = comandoORA.execute
                   Set comandoORA = Nothing
                   Set parametro = Nothing
                   var_Titular_vianney_catalog = rsaux9!vcha_tit_titular_id
                   VAR_ESTABLECIMIENTO = rsaux9!vcha_esb_establecimient_id
                   var_si_vianney_Catalog = 0
                   var_almacen_Destino = rsaux9!vcha_esb_establecimient_id
                   If var_Titular_vianney_catalog = "T000000343" Then
                      var_si_vianney_Catalog = 1
                   End If
                   var_dia = Day(Date)
                   var_mes = Month(Date)
                   var_año = Year(Date)
                         
                   If var_mes = 1 Then
                      var_mes_str = "Enero"
                   End If
                   If var_mes = 2 Then
                      var_mes_str = "Febrero"
                   End If
                   If var_mes = 3 Then
                      var_mes_str = "Marzo"
                   End If
                   If var_mes = 4 Then
                      var_mes_str = "Abril"
                   End If
                   If var_mes = 5 Then
                      var_mes_str = "Mayo"
                   End If
                   If var_mes = 6 Then
                      var_mes_str = "Junio"
                   End If
                   If var_mes = 7 Then
                      var_mes_str = "Julio"
                   End If
                   If var_mes = 8 Then
                      var_mes_str = "Agosto"
                   End If
                   If var_mes = 10 Then
                      var_mes_str = "Octubre"
                   End If
                   If var_mes = 11 Then
                      var_mes_str = "Noviembre"
                   End If
                   If var_mes = 12 Then
                      var_mes_str = "Diciembre"
                   End If
                         
                   var_fecha_str_carta = CStr(var_dia) + " de " + var_mes_str + " del " + CStr(var_año)
                   If var_si_vianney_Catalog = 1 Then
                      strconsulta = "SELECT * FROM RA_CUSTOMER_tRX_ALL WHERE CT_REFERENCE = ? "
                      With comandoORA
                           .ActiveConnection = cnnoracle_4
                           .CommandType = adCmdText
                           .CommandText = strconsulta
                           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CDbl(rsaux12!source_header_number))
                           .Parameters.Append parametro
                      End With
                      Set rsaux1 = comandoORA.execute
                      Set comandoORA = Nothing
                      Set parametro = Nothing
                      If Not rsaux1.EOF Then
                         var_dia = Day(rsaux1!trx_date)
                         var_mes = Month(rsaux1!trx_date)
                         var_año = Year(rsaux1!trx_date)
                         
                         If var_mes = 1 Then
                            var_mes_str = "Enero"
                         End If
                         If var_mes = 2 Then
                            var_mes_str = "Febrero"
                         End If
                         If var_mes = 3 Then
                            var_mes_str = "Marzo"
                         End If
                         If var_mes = 4 Then
                            var_mes_str = "Abril"
                         End If
                         If var_mes = 5 Then
                            var_mes_str = "Mayo"
                         End If
                         If var_mes = 6 Then
                            var_mes_str = "Junio"
                         End If
                         If var_mes = 7 Then
                            var_mes_str = "Julio"
                         End If
                         If var_mes = 8 Then
                            var_mes_str = "Agosto"
                         End If
                         If var_mes = 10 Then
                            var_mes_str = "Octubre"
                         End If
                         If var_mes = 11 Then
                            var_mes_str = "Noviembre"
                         End If
                         If var_mes = 12 Then
                            var_mes_str = "Diciembre"
                         End If
                         
                         var_fecha_str = " con fecha del " + CStr(var_dia) + " de " + var_mes_str + " del " + CStr(var_año)
                         
                         If var_cadena_facturas = "" Then
                            var_cadena_facturas = "FAEVII " + CStr(rsaux1!TRX_NUMBER) + var_fecha_str
                         Else
                            var_cadena_facturas = var_cadena_facturas + ", FAEVII " + CStr(rsaux1!TRX_NUMBER) + var_fecha_str
                         End If
                      Else
                         VAR_TODAS_LAS_FACTURAS = 0
                      End If
                      rsaux1.Close
                   End If
                   rsaux12.MoveNext
             Wend
             rsaux12.Close
          End If
      Next var_j
      If VAR_TODAS_LAS_FACTURAS = 1 Then
         cnn.BeginTrans
         rs.Open "SELECT MAX(INTE_TEM_cONSECUTIVO) FROM TB_TEMP_ORACLE_CARTA_JURAMENTADA", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rs.Close
         rs.Open "INSERT INTO TB_TEMP_ORACLE_CARTA_JURAMENTADA (INTE_TEM_CONSECUTIVO, FACTURAS, fecha) VALUES (" + CStr(var_consecutivo) + ",'" + var_cadena_facturas + "','Aguascalientes, Ags; a " + var_fecha_str_carta + "')", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_carta_juramentada.rpt")
                  reporte.RecordSelectionFormula = "{tb_temp_oracle_carta_juramentada.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  'reporte.ExportOptions.FormatType = crEFTPortableDocFormat
                  reporte.ExportOptions.FormatType = crEFTWordForWindows
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\carta_juramentada_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.doc"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  Me.txt_embarque = ""
                  MsgBox "Se a terminado de guardar el archivo " + archivo
      Else
         MsgBox "Faltan pedidos por facturar", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se han selecciionado embarques", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2100
   Left = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         var_cadena = "SELECT * from xxvia_tb_encabezado_embarques where embarque = ?"
         strconsulta = var_cadena
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_embarque)
              .Parameters.Append parametro
         End With
         Set rsaux12 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux12.EOF Then
            VAR_ESTATUS = IIf(IsNull(rsaux12!CHAR_EMB_ESTATUS), "", rsaux12!CHAR_EMB_ESTATUS)
            If VAR_ESTATUS = "I" Or VAR_ESTATUS = "F" Then
               If Me.lv_lista.ListItems.Count > 0 Then
                  var_posible = 1
                  For var_j = 1 To Me.lv_lista.ListItems.Count
                      Me.lv_lista.ListItems.Item(var_j).Selected = True
                      If CDbl(Me.lv_lista.selectedItem) = CDbl(Me.txt_embarque) Then
                         var_posible = 0
                      End If
                  Next var_j
                  If var_posible = 1 Then
                     Set list_item = lv_lista.ListItems.Add(, , Me.txt_embarque)
                     Me.txt_embarque = ""
                  End If
               Else
                  Set list_item = lv_lista.ListItems.Add(, , Me.txt_embarque)
                  Me.txt_embarque = ""
               End If
            Else
               MsgBox "El embarque aun no a sido cerrado", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rsaux12.Close
      End If
   End If
End Sub

