VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_audita_caja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auditar caja"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   165
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   90
      TabIndex        =   6
      Top             =   5520
      Width           =   2325
   End
   Begin VB.TextBox txt_total 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3360
      TabIndex        =   4
      Top             =   5535
      Width           =   3045
   End
   Begin VB.Frame Frame2 
      Height          =   1125
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   6300
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   300
         TabIndex        =   2
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4185
      Left            =   90
      TabIndex        =   0
      Top             =   1290
      Width           =   6315
      Begin MSComctlLib.ListView lv_lista 
         Height          =   3945
         Left            =   45
         TabIndex        =   3
         Top             =   150
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   6959
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6262
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   630
      URL             =   "C:\sistemas\desarrollo\integral\type.wma"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1111
      _cy             =   238
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp_1 
      Height          =   795
      Left            =   210
      TabIndex        =   7
      Top             =   1800
      Width           =   0
      URL             =   "\\tsclient\C\sistemas\desarrollo\oracle\000867298_prev.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   3
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   0
      _cy             =   1402
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2610
      TabIndex        =   5
      Top             =   5610
      Width           =   690
   End
End
Attribute VB_Name = "frmoracle_audita_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
   Unload Me
End Sub

Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.play
End Sub

Private Sub Form_Load()
   Me.wmp_1.URL = App.Path + "\Mec_Alarm_10.wav"
   rsaux2.Open "DELETE FROM XXVIA_TB_CAJAS_AUDITADAS WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar), cnnoracle_4, adOpenDynamic, adLockOptimistic
   rsaux2.Open "SELECT segment1, sum(FLOA_SAL_CANTIDAD_LEIDA) as FLOA_SAL_CANTIDAD_LEIDA FROM XXVIA_tB_sALIDAS_CAJAS WHERE INTE_eMB_EMBARQUE = " + CStr(var_embarque_auditar) + " AND INTE_PAQ_CAJA = " + CStr(var_caja_auditar) + " group by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rsaux2.EOF
         rsaux4.Open "INSERT INTO XXVIA_TB_CAJAS_AUDITADAS (EMBARQUE, CAJA, CODIGO, CANTIDAD_ORIGINAL, CANTIDAD_AUDITADA) VALUES (" + CStr(var_embarque_auditar) + "," + CStr(var_caja_auditar) + ",'" + rsaux2!SEGMENT1 + "'," + CStr(rsaux2!FLOA_SAL_CANTIDAD_LEIDA) + ",0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux2.MoveNext
   Wend
   rsaux2.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If rsaux9.State = 1 Then
      rsaux9.Close
   End If
End Sub

Private Sub lv_lista_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      rs.Open "UPDATE XXVIA_TB_CAJAS_AUDITADAS SET CANTIDAD_AUDITADA = CANTIDAD_AUDITADA - 1 WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CODIGO = '" + Me.lv_lista.selectedItem + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Me.lv_lista.selectedItem.SubItems(2) = CDbl(Me.lv_lista.selectedItem.SubItems(2)) - 1
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim clnt As New SoapClient30
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
   If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      var_localizador_subinventario = " "
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
      rsaux9.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If rsaux9.EOF Then
         rsaux8.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux8.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, b.UNIT_WEIGHT FROM mtl_cross_references_v A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
            If IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador) <> "" Then
               var_localizador_subinventario = txt_almacen + IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador)
               If var_localizador_subinventario <> "" Then
                  Me.txt_codigo = rsaux8!SEGMENT1
               Else
                  Me.txt_codigo = ""
               End If
            Else
               Me.txt_codigo = ""
            End If
         Else
            Me.txt_codigo = ""
         End If
         rsaux8.Close
      End If
      rsaux9.Close
      If Me.txt_codigo <> "" Then
         rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
            var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
            var_encontro = 0
            var_cantidad = 1
            For var_j = 1 To lv_lista.ListItems.Count
                lv_lista.ListItems.Item(var_j).Selected = True
                If lv_lista.selectedItem = Me.txt_codigo Then
                   var_encontro = var_j
                End If
            Next var_j
            var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
            If var_salida_masiva = "Y" Then
               var_codigo_global = Me.txt_codigo
               frmoracle_cantidad.Show 1
               var_cantidad_leida = var_cantidad_global
               Me.txt_codigo = var_codigo_global
            Else
               var_cantidad_leida = 1
            End If
            
            If var_encontro > 0 Then
               lv_lista.ListItems.Item(var_encontro).Selected = True
               lv_lista.selectedItem.SubItems(2) = Format(CDbl(lv_lista.selectedItem.SubItems(2)) + var_cantidad_leida, "###,###,##0.00")
            Else
               Set list_item = Me.lv_lista.ListItems.Add(, , Me.txt_codigo)
               list_item.SubItems(1) = rsaux8!Description
               list_item.SubItems(2) = Format(var_cantidad_leida, "###,###,##0.00")
            End If
            rsaux2.Open "SELECT * FROM XXVIA_TB_cAJAS_AUDITADAS WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CODIGO = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               rsaux.Open "UPDATE XXVIA_TB_CAJAS_AUDITADAS SET CANTIDAD_AUDITADA = CANTIDAD_AUDITADA + " + CStr(var_cantidad_leida) + "  WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CODIGO = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               Call cmd_mensaje_4_Click
               If var_modo_texto_ip = 1 Then
                  On Error GoTo SALIR:
                  Set clnt = Nothing
                  clnt.MSSoapInit var_webservice_texto
                  var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", AUDITOR: " + var_nombre_usuario + Chr(13) + CStr(var_embarque_auditar) + "-" + CStr(var_caja_auditar) + "-" + Me.txt_codigo + "   " + rsaux8!Description + " CANTIDAD: " + CStr(var_cantidad_leida) + Chr(13))
                  Set clnt = Nothing
               End If
            
            Else
               rsaux.Open "INSERT INTO XXVIA_TB_CAJAS_AUDITADAS (EMBARQUE, CAJA, CODIGO, CANTIDAD_ORIGINAL, CANTIDAD_AUDITADA) VALUES (" + CStr(var_embarque_auditar) + "," + CStr(var_caja_auditar) + ",'" + Me.txt_codigo + "',0," + CStr(var_cantidad_leida) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
               Call cmd_mensaje_4_Click
               If var_modo_texto_ip = 1 Then
                  On Error GoTo SALIR:
                  Set clnt = Nothing
                  clnt.MSSoapInit var_webservice_texto
                  var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", AUDITOR: " + var_nombre_usuario + Chr(13) + CStr(var_embarque_auditar) + "-" + CStr(var_caja_auditar) + "-" + Me.txt_codigo + "   " + rsaux8!Description + " CANTIDAD: " + CStr(var_cantidad_leida) + Chr(13))
                  Set clnt = Nothing
               End If
            End If
            rsaux2.Close
            Me.txt_codigo = ""
            var_total = 0
            For var_j = 1 To lv_lista.ListItems.Count
                lv_lista.ListItems.Item(var_j).Selected = True
                var_total = var_total + lv_lista.selectedItem.SubItems(2)
            Next var_j
            Me.txt_total = Format(var_total, "###,###,##0.00")
            
         Else
            
            txt_codigo = ""
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "El artículo no existe"
            frmmensaje.Show 1
         End If
         rsaux8.Close
      Else
         txt_codigo = ""
         frmmensaje.lbl_articulo = ""
         frmmensaje.lbl_mensaje = "El artículo no existe"
         frmmensaje.Show 1
      End If
   End If
   Exit Sub
SALIR:
    Resume Next
End Sub
