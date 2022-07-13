VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcarta_porte_traspasos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta porte traspasos"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_folio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   4680
      Width           =   3135
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmcarta_porte_traspasos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "frmcarta_porte_traspasos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   3000
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   8500
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2415
         Left            =   45
         TabIndex        =   4
         Top             =   375
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4260
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10080
      Picture         =   "frmcarta_porte_traspasos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   120
      TabIndex        =   7
      Top             =   330
      Width           =   10365
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   510
      Picture         =   "frmcarta_porte_traspasos.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   5115
      Left            =   120
      TabIndex        =   9
      Top             =   405
      Width           =   10395
      Begin VB.TextBox txt_chofer 
         Height          =   390
         Left            =   990
         TabIndex        =   27
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre_chofer 
         Height          =   390
         Left            =   2820
         TabIndex        =   26
         Top             =   600
         Width           =   6975
      End
      Begin VB.TextBox txt_embarque 
         Height          =   390
         Left            =   990
         TabIndex        =   25
         Top             =   165
         Width           =   1815
      End
      Begin VB.TextBox txt_unidad 
         Height          =   390
         Left            =   1005
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre_unidad 
         Height          =   390
         Left            =   2880
         TabIndex        =   23
         Top             =   1680
         Width           =   6975
      End
      Begin VB.TextBox txt_RFC 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1005
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_licencia 
         Enabled         =   0   'False
         Height          =   390
         Left            =   3885
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_permsct 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1005
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_numpermisosct 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4725
         TabIndex        =   19
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_seguro 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1005
         TabIndex        =   18
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox txt_poliza 
         Enabled         =   0   'False
         Height          =   390
         Left            =   6045
         TabIndex        =   17
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txt_configuracion_vehicular 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1920
         TabIndex        =   16
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_placaVM 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4605
         TabIndex        =   15
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_modelo_VM 
         Enabled         =   0   'False
         Height          =   390
         Left            =   7605
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_remolque 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1920
         TabIndex        =   13
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txt_placa_remolque 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4605
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Height          =   75
         Left            =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   10365
      End
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   0
         TabIndex        =   10
         Top             =   4080
         Width           =   10365
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Folio: "
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   4440
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Chofer"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1785
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Licencia:"
         Height          =   195
         Left            =   3120
         TabIndex        =   38
         Top             =   1185
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Perm. SCT:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número permiso SCT:"
         Height          =   195
         Left            =   3120
         TabIndex        =   36
         Top             =   2160
         Width           =   1545
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Seguro:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2745
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Poliza:"
         Height          =   195
         Left            =   5040
         TabIndex        =   34
         Top             =   2745
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Configuración vehicular:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   3225
         Width           =   1710
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Placa VM:"
         Height          =   195
         Left            =   3840
         TabIndex        =   32
         Top             =   3225
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Modelo VM:"
         Height          =   195
         Left            =   6720
         TabIndex        =   31
         Top             =   3225
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tipo remolque:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   3705
         Width           =   1050
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Número permiso SCT:"
         Height          =   195
         Left            =   3000
         TabIndex        =   29
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Left            =   3840
         TabIndex        =   28
         Top             =   3705
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chofer"
      Height          =   195
      Left            =   240
      TabIndex        =   44
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Unidad:"
      Height          =   195
      Left            =   3240
      TabIndex        =   43
      Top             =   2640
      Width           =   75
   End
End
Attribute VB_Name = "frmcarta_porte_traspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo As Integer
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection

Private Sub cmd_imprimir_Click()
      Dim var_serie As String
      Dim var_folio As String
      If Me.txt_folio <> "" Then
         var_cadena = "SELECT cod.segment1,cod.description, tr.organization_id, case when tr.SUBINVENTORY_CODE ='ALMCAJAS' then tr.transaction_source_name else tr.SUBINVENTORY_CODE end almacen_origen, case when tr.transaction_type_id = 2 then case when tr.TRANSACTION_SOURCE_NAME is null then tr.TRANSFER_SUBINVENTORY else tr.TRANSACTION_SOURCE_NAME end else tr.TRANSFER_SUBINVENTORY end almacen_destino, primary_quantity * -1 cantidad, nvl(cod.CLASIFICACIONSAT,'01010101'), tr.last_update_date, cod.UOM_SAT, PRIMARY_UNIT_OF_MEASURE AS UNIDAD_MEDIDA , tr.SUBINVENTORY_CODE || '-' || case when tr.transaction_type_id = 2 then 3 else 21 end Serie, SUBSTR( tr.SHIPMENT_NUMBER,LASTINDEXOF(tr.SHIPMENT_NUMBER,'-') +1 ) Folio, TR.TRANSFER_ORGANIZATION_ID AS ORGANIZACION_DESTINO,(select sum(primary_quantity * -1) From mtl_material_transactions"
         var_cadena = var_cadena + " where tr.TRANSACTION_TYPE_ID  in( 2, 21) and decode(TRANSACTION_TYPE_ID,21,'TRANS', decode( TRANSACTION_SOURCE_NAME, null, 'TRANS',transfer_subinventory)) LIKE '%TRANS%' AND shipment_number = ? AND primary_quantity      < 0) pzas_total FROM mtl_material_transactions tr, xxvia_system_items_b cod,      INV.MTL_SECONDARY_INVENTORIES sinO, INV.MTL_SECONDARY_INVENTORIES sinD WHERE tr.TRANSACTION_TYPE_ID  in( 2, 21) and tr.ORGANIZATION_ID = ? and decode(tr.TRANSACTION_TYPE_ID,21,'TRANS', decode( tr.TRANSACTION_SOURCE_NAME, null, 'TRANS',tr.transfer_subinventory)) LIKE '%TRANS%' AND tr.shipment_number    = ? AND tr.primary_quantity      < 0 AND cod.inventory_item_id    = tr.inventory_item_id AND cod.organization_id      = tr.organization_id and tr.ORGANIZATION_ID = sinO.ORGANIZATION_ID and tr.SUBINVENTORY_CODE = sinO.SECONDARY_INVENTORY_NAME"
         var_cadena = var_cadena + " and tr.TRANSFER_ORGANIZATION_ID = sinD.ORGANIZATION_ID and decode( tr.transaction_type_id, 2, case when tr.TRANSACTION_SOURCE_NAME is null then tr.TRANSFER_SUBINVENTORY else tr.TRANSACTION_SOURCE_NAME end,tr.TRANSFER_SUBINVENTORY) = decode(tr.SUBINVENTORY_CODE, 'ALMCAJAS',case when tr.TRANSACTION_SOURCE_NAME is null then tr.TRANSFER_SUBINVENTORY else tr.TRANSACTION_SOURCE_NAME end, sinD.SECONDARY_INVENTORY_NAME)"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = var_cadena
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         If Not rs.EOF Then
            var_serie = rs!Serie
            var_folio = rs!Folio
            var_tipo = 2
            var_cadena = "CALL XXVIA_SP_TIMBRAR_TRASPASOS_12(?,?,?)"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = var_cadena
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_folio)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                .Parameters.Append parametro
            End With
            Set rsaux11 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
                   
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            var_cadena = "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'CPTR" + Me.txt_folio + "' and numero = " + CStr(var_folio)
            rsaux1.Open "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'CPTR" + Me.txt_folio + "_' and numero = " + CStr(var_folio), cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               var_cadena = Replace(Replace(rsaux1!Cadena, "T23:", "T00:"), "AUTORIZADO  ", "AUTORIZADO ")
               var_cadena_rfc = Mid(var_cadena, 34, 12)
               VAR_CADENA_STR = ""
               Open ("C:\SISTEMAS\CPTR" + Trim(var_serie) + "_" + Trim(Str(var_folio)) + ".FAC") For Output As #1
               For var_i = 1 To Len(var_cadena)
                   If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                      Print #1, VAR_CADENA_STR
                      VAR_CADENA_STR = ""
                   Else
                      VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                   End If
               Next var_i
               Print #1, "FIN:"
               Close #1
                        
               var_archivo = "C:\SISTEMAS\sube_fact_" + Trim("CPTR" + var_serie) + "_" + Trim(Str(VAAR_FOLIO)) + ".bat"
               x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\CPTR" + Trim(var_serie) + "_" + var_folio + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
            Else
               MsgBox "El traspaso no existe", vbOKOnly, "ATENCION"
            End If
            rsaux1.Close
         Else
            MsgBox "El folio no existe.", vbOKOnly, "ATENCION"
         End If
      Else
      End If

End Sub

Private Sub cmd_nuevo_Click()
   var_si = MsgBox("¿Desea crear un embarque?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "select MAX(EMBARQUE) AS MAXIMO_EMBARQUE from XXVIA_TB_ENCABEZADO_EMBARQUES", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         var_numero_embarque = 1
      Else
         var_numero_embarque = IIf(IsNull(rs!maximo_embarque), 0, rs!maximo_embarque) + 1
      End If
      rs.Close
      Me.txt_embarque = var_numero_embarque
      var_cadena = "insert into xxvia_tb_encabezado_embarques (EMBARQUE) "
      var_cadena = var_cadena + " values (" + CStr(var_numero_embarque) + ")"
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      Me.txt_chofer.SetFocus
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   If IsNumeric(Me.txt_embarque) Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If Me.txt_chofer <> "" Then
            If Me.txt_unidad <> "" Then
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHOFER_cdmx = '" + Me.txt_chofer + "', TRANSPORTE_cdmx = '" + Me.txt_unidad + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "Unidad invalida.", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Chofer invalido.", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque " + Me.txt_embarque + " no existe.", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de embarque invalido.", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If var_tipo = 1 Then
          Me.txt_chofer = Me.lv_lista.selectedItem
          Me.txt_nombre_chofer = Me.lv_lista.selectedItem.SubItems(1)
          rs.Open "select * from xxvia_tb_choferes where id_chofer = '" + Me.lv_lista.selectedItem + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             Me.txt_RFC = IIf(IsNull(rs!rfc), "", rs!rfc)
             Me.txt_licencia = IIf(IsNull(rs!licencia), "", rs!licencia)
          End If
          Me.txt_chofer.SetFocus
          rs.Close
       End If
       If var_tipo = 2 Then
          Me.txt_unidad = Me.lv_lista.selectedItem
          Me.txt_nombre_unidad = Me.lv_lista.selectedItem.SubItems(1)
          rsaux.Open "SELECT * FROM XXVIA_tB_tRANSPORTES WHERE CLAVE = '" + txt_unidad + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rsaux.EOF Then
             Me.txt_permsct = IIf(IsNull(rsaux!PERMSCT), "", rsaux!PERMSCT)
             Me.txt_seguro = IIf(IsNull(rsaux!NOMBREASEG), "", rsaux!NOMBREASEG)
             Me.txt_numpermisosct = IIf(IsNull(rsaux!NUMPERMIsoSCT), "", rsaux!NUMPERMIsoSCT)
             Me.txt_poliza = IIf(IsNull(rsaux!NUMPOLIZASEG), "", rsaux!NUMPOLIZASEG)
             Me.txt_configuracion_vehicular = IIf(IsNull(rsaux!configvehicular), "", rsaux!configvehicular)
             Me.txt_placaVM = IIf(IsNull(rsaux!placavm), "", rsaux!placavm)
             Me.txt_modelo_VM = IIf(IsNull(rsaux!aniomodelovm), "", rsaux!aniomodelovm)
             Me.txt_remolque = IIf(IsNull(rsaux!subtiporem), "", rsaux!subtiporem)
             Me.txt_placa_remolque = IIf(IsNull(rsaux!placas), "", rsaux!placas)
          End If
          rsaux.Close
          Me.txt_unidad.SetFocus
       End If
    End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_chofer_Change()
   Me.txt_nombre_chofer = ""
   Me.txt_RFC = ""
   Me.txt_licencia = ""
End Sub

Private Sub txt_chofer_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo = 1
      Me.frm_lista.Visible = True
      Me.lv_lista.ListItems.Clear
      rs.Open "select * from xxvia_tb_choferes where nvl(rfc,' ')<> ' ' order by nombre", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!id_chofer)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      Me.lv_lista.SetFocus
   End If

End Sub

Private Sub txt_chofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_unidad.SetFocus
   End If
End Sub

Private Sub txt_embarque_Change()
   Me.txt_chofer = ""
   Me.txt_configuracion_vehicular = ""
   Me.txt_licencia = ""
   Me.txt_modelo_VM = ""
   Me.txt_nombre_chofer = ""
   Me.txt_nombre_unidad = ""
   Me.txt_numpermisosct = ""
   Me.txt_permsct = ""
   Me.txt_placa_remolque = ""
   Me.txt_placaVM = ""
   Me.txt_poliza = ""
   Me.txt_remolque = ""
   Me.txt_RFC = ""
   Me.txt_seguro = ""
   Me.txt_unidad = ""
   Me.txt_folio = ""
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "select * from xxvia_Tb_encabezado_embarques where embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_chofer = IIf(IsNull(rs!CHOFER_CDMX), "", rs!CHOFER_CDMX)
            If var_chofer <> "" Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "SELECT * FROM XXVIA_TB_CHOFERES WHERE ID_CHOFER = '" + var_chofer + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  Me.txt_chofer = IIf(IsNull(rsaux!id_chofer), "", rsaux!id_chofer)
                  Me.txt_nombre_chofer = IIf(IsNull(rsaux!NOMBRE), "", rsaux!NOMBRE)
                  Me.txt_licencia = IIf(IsNull(rsaux!licencia), "", rsaux!licencia)
                  Me.txt_RFC = IIf(IsNull(rsaux!rfc), "", rsaux!rfc)
               End If
               rsaux.Close
            Else
               MsgBox "El embarque no tiene chofer asignado.", vbOKOnly, "ATENCION"
            End If
            var_transporte = IIf(IsNull(rs!transporte_CDMX), "", rs!transporte_CDMX)
            If var_transporte <> "" Then
               rsaux.Open "SELECT * FROM XXVIA_tB_tRANSPORTES WHERE CLAVE = '" + var_transporte + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  Me.txt_unidad = IIf(IsNull(rsaux!clave), "", rsaux!clave)
                  Me.txt_nombre_unidad = IIf(IsNull(rsaux!NOMBRE), "", rsaux!NOMBRE)
                  Me.txt_permsct = IIf(IsNull(rsaux!PERMSCT), "", rsaux!PERMSCT)
                  Me.txt_seguro = IIf(IsNull(rsaux!NOMBREASEG), "", rsaux!NOMBREASEG)
                  Me.txt_numpermisosct = IIf(IsNull(rsaux!NUMPERMIsoSCT), "", rsaux!NUMPERMIsoSCT)
                  Me.txt_poliza = IIf(IsNull(rsaux!NUMPOLIZASEG), "", rsaux!NUMPOLIZASEG)
                  Me.txt_configuracion_vehicular = IIf(IsNull(rsaux!configvehicular), "", rsaux!configvehicular)
                  Me.txt_placaVM = IIf(IsNull(rsaux!placavm), "", rsaux!placavm)
                  Me.txt_modelo_VM = IIf(IsNull(rsaux!aniomodelovm), "", rsaux!aniomodelovm)
                  Me.txt_remolque = IIf(IsNull(rsaux!subtiporem), "", rsaux!subtiporem)
                  Me.txt_placa_remolque = IIf(IsNull(rsaux!placas), "", rsaux!placas)
                  
               End If
               rsaux.Close
            Else
               MsgBox "El embarque no tiene transporte asignado.", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.txt_chofer.SetFocus
      Else
         MsgBox "Número de embarque incorrecto.", vbOKOnly, "ATENCION"
      End If
   End If

End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_serie As String
      Dim var_folio As String
      If Me.txt_folio <> "" Then
         var_cadena = "SELECT cod.segment1,cod.description, tr.organization_id, case when tr.SUBINVENTORY_CODE ='ALMCAJAS' then tr.transaction_source_name else tr.SUBINVENTORY_CODE end almacen_origen, case when tr.transaction_type_id = 2 then case when tr.TRANSACTION_SOURCE_NAME is null then tr.TRANSFER_SUBINVENTORY else tr.TRANSACTION_SOURCE_NAME end else tr.TRANSFER_SUBINVENTORY end almacen_destino, primary_quantity * -1 cantidad, nvl(cod.CLASIFICACIONSAT,'01010101'), tr.last_update_date, cod.UOM_SAT, PRIMARY_UNIT_OF_MEASURE AS UNIDAD_MEDIDA , tr.SUBINVENTORY_CODE || '-' || case when tr.transaction_type_id = 2 then 3 else 21 end Serie, SUBSTR( tr.SHIPMENT_NUMBER,LASTINDEXOF(tr.SHIPMENT_NUMBER,'-') +1 ) Folio, TR.TRANSFER_ORGANIZATION_ID AS ORGANIZACION_DESTINO,(select sum(primary_quantity * -1) From mtl_material_transactions"
         var_cadena = var_cadena + " where tr.TRANSACTION_TYPE_ID  in( 2, 21) and decode(TRANSACTION_TYPE_ID,21,'TRANS', decode( TRANSACTION_SOURCE_NAME, null, 'TRANS',transfer_subinventory)) LIKE '%TRANS%' AND shipment_number = ? AND primary_quantity      < 0) pzas_total FROM mtl_material_transactions tr, xxvia_system_items_b cod,      INV.MTL_SECONDARY_INVENTORIES sinO, INV.MTL_SECONDARY_INVENTORIES sinD WHERE tr.TRANSACTION_TYPE_ID  in( 2, 21) and tr.ORGANIZATION_ID = ? and decode(tr.TRANSACTION_TYPE_ID,21,'TRANS', decode( tr.TRANSACTION_SOURCE_NAME, null, 'TRANS',tr.transfer_subinventory)) LIKE '%TRANS%' AND tr.shipment_number    = ? AND tr.primary_quantity      < 0 AND cod.inventory_item_id    = tr.inventory_item_id AND cod.organization_id      = tr.organization_id and tr.ORGANIZATION_ID = sinO.ORGANIZATION_ID and tr.SUBINVENTORY_CODE = sinO.SECONDARY_INVENTORY_NAME"
         var_cadena = var_cadena + " and tr.TRANSFER_ORGANIZATION_ID = sinD.ORGANIZATION_ID and decode( tr.transaction_type_id, 2, case when tr.TRANSACTION_SOURCE_NAME is null then tr.TRANSFER_SUBINVENTORY else tr.TRANSACTION_SOURCE_NAME end,tr.TRANSFER_SUBINVENTORY) = decode(tr.SUBINVENTORY_CODE, 'ALMCAJAS',case when tr.TRANSACTION_SOURCE_NAME is null then tr.TRANSFER_SUBINVENTORY else tr.TRANSACTION_SOURCE_NAME end, sinD.SECONDARY_INVENTORY_NAME)"

         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = var_cadena
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(Me.txt_folio))
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
            Me.cmd_imprimir.SetFocus
         Else
            MsgBox "El folio no existe.", vbOKOnly, "ATENCION"
            Me.txt_folio = ""
         End If
         rs.Close
      Else
      End If
   End If

End Sub

Private Sub txt_unidad_Change()
   Me.txt_configuracion_vehicular = ""
   Me.txt_modelo_VM = ""
   Me.txt_nombre_unidad = ""
   Me.txt_numpermisosct = ""
   Me.txt_permsct = ""
   Me.txt_placa_remolque = ""
   Me.txt_placaVM = ""
   Me.txt_poliza = ""
   Me.txt_remolque = ""
   Me.txt_seguro = ""

End Sub

Private Sub txt_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo = 2
      Me.frm_lista.Visible = True
      Me.lv_lista.ListItems.Clear
      rs.Open "select * from xxvia_tb_transportes order by nombre", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      Me.lv_lista.SetFocus
      
   End If

End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_folio.SetFocus
   End If
End Sub
