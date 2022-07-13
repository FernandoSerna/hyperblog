VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcodigos_master 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Códigos master"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1740
      Picture         =   "frmcodigos_master.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir reporte"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cargar_pedido 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Picture         =   "frmcodigos_master.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cargar códigos master"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmcodigos_master.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprimir etiqueta"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmcodigos_master.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmcodigos_master.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmcodigos_master.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9255
      Picture         =   "frmcodigos_master.frx":060C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   105
      TabIndex        =   13
      Top             =   270
      Width           =   9555
   End
   Begin VB.Frame Frame2 
      Height          =   1620
      Left            =   120
      TabIndex        =   9
      Top             =   345
      Width           =   9540
      Begin VB.TextBox txt_cantidad 
         Height          =   375
         Left            =   915
         TabIndex        =   7
         Top             =   1110
         Width           =   1410
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   375
         Left            =   2295
         TabIndex        =   6
         Top             =   705
         Width           =   7170
      End
      Begin VB.TextBox txt_codigo 
         Height          =   375
         Left            =   915
         TabIndex        =   5
         Top             =   705
         Width           =   1365
      End
      Begin VB.TextBox txt_master 
         Height          =   375
         Left            =   915
         TabIndex        =   4
         Top             =   300
         Width           =   2985
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   795
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Master:"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   390
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4950
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   9540
      Begin MSComctlLib.ListView lv_etiquetas 
         Height          =   4755
         Left            =   15
         TabIndex        =   14
         Top             =   135
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   8387
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Master"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
   End
End
Attribute VB_Name = "frmcodigos_master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_cargar_pedido_Click()
   var_importar_codigos_master_de = 0
   frmimportar_codigos_master_de.Show 1
   If var_empresa = "28" Then
      If var_importar_codigos_master_de = 1 Then
         On Error GoTo salir
         strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb); DBQ=c:\SISTEMAS\master.xls"
         rsaux2.Open "SELECT * FROM [master$]", strConnectionString
         While Not rsaux2.EOF
               rsaux3.Open "select * from tb_articulos where vcha_art_Articulo_id = '" + IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  rs.Open "SELECT * FROM TB_ETIQUETAS_MASTER WHERE VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux2!MASTER), "", rsaux2!MASTER)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "update TB_ETIQUETAS_MASTER set floa_eti_Cantidad = " + CStr(IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)) + " where VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux2!MASTER), "", rsaux2!MASTER)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO) + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "insert into tb_etiquetas_master (vcha_eti_etiqueta_master, vcha_art_articulo_id, floa_eti_cantidad) values ('" + CStr(IIf(IsNull(rsaux2!MASTER), "", rsaux2!MASTER)) + "', '" + UCase(IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO)) + "', " + CStr(IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
               End If
               rsaux3.Close
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         Me.lv_etiquetas.ListItems.Clear
         rs.Open "select vcha_eti_etiqueta_master, a.vcha_art_articulo_id, vcha_art_nombre_español, floa_eti_cantidad from tb_etiquetas_master a, TB_ARTICULOS b where  a.vcha_art_articulo_id = b.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_etiquetas.ListItems.Add(, , rs!vcha_eti_etiqueta_master)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id)
               list_item.SubItems(2) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ETI_CANTIDAD), "", rs!FLOA_ETI_CANTIDAD)
               rs.MoveNext
         Wend
         rs.Close
      End If
      If var_importar_codigos_master_de = 2 Then
         var_cadena = "select fac.*, lec.VCHA_CAJ_CAJA_ID from XXVIA_VW_FACTURAS_CANTIA fac, XXVIA.XXVIA_TB_LEC_MOV_INV lec where fac.TRX_NUMBER = '" + var_numero_factura + "'  and fac.CUSTOMER_TRX_ID = lec.VCHA_LMI_LOCALIZADOR and fac.SEGMENT1 = lec.VCHA_LMI_CODIGO AND SERIE = '" + var_serie + "'"
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  rsaux2.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE substring(VCHA_EQU_CODIGO_EQUIVALENTE,1,8) = '" + rsaux!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     If Mid(rsaux!segment1, 1, 2) = "00" Then
                        If Len(IIf(IsNull(rsaux2!vcha_art_articulo_id), "", rsaux2!vcha_art_articulo_id)) <= 8 Then
                           'VER LO DE LAS EQUIVALENCIAS
                           'var_articulo_enviado = Mid(IIf(IsNull(rsaux2!vcha_Art_Articulo_id), "", rsaux2!vcha_Art_Articulo_id), 1, 5)
                           'Else
                           var_articulo_enviado = IIf(IsNull(rsaux2!vcha_art_articulo_id), "", rsaux2!vcha_art_articulo_id)
                        End If
                     Else
                        var_articulo_enviado = IIf(IsNull(rsaux2!vcha_art_articulo_id), "", rsaux2!vcha_art_articulo_id)
                     End If
                  End If
                  rsaux2.Close
                     
                  If var_articulo_enviado <> "" Then
                     rsaux11.Open "SELECT * FROM TB_ETIQUETAS_MASTER WHERE VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux!VCHA_CAJ_CAJA_ID), "", rsaux!VCHA_CAJ_CAJA_ID)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(var_articulo_enviado), "", var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux11.EOF Then
                        rsaux9.Open "update TB_ETIQUETAS_MASTER set floa_eti_Cantidad = " + CStr(IIf(IsNull(rsaux!NUMB_LMI_CANTIDAD), 0, rsaux!NUMB_LMI_CANTIDAD)) + " where VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux!VCHA_CAJ_CAJA_ID), "", rsaux!VCHA_CAJ_CAJA_ID)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(var_articulo_enviado), "", var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux9.Open "insert into tb_etiquetas_master (vcha_eti_etiqueta_master, vcha_art_articulo_id, floa_eti_cantidad) values ('" + CStr(IIf(IsNull(rsaux!VCHA_CAJ_CAJA_ID), "", rsaux!VCHA_CAJ_CAJA_ID)) + "', '" + UCase(IIf(IsNull(var_articulo_enviado), "", var_articulo_enviado)) + "', " + CStr(IIf(IsNull(rsaux!NUMB_LMI_CANTIDAD), 0, rsaux!NUMB_LMI_CANTIDAD)) + ")", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux11.Close
                  End If
                  rsaux.MoveNext
            Wend
         End If
         
         
         
         
         
         Me.lv_etiquetas.ListItems.Clear
         rs.Open "select vcha_eti_etiqueta_master, a.vcha_art_articulo_id, vcha_art_nombre_español, floa_eti_cantidad from tb_etiquetas_master a, TB_ARTICULOS b where  a.vcha_art_articulo_id = b.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_etiquetas.ListItems.Add(, , rs!vcha_eti_etiqueta_master)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id)
               list_item.SubItems(2) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ETI_CANTIDAD), "", rs!FLOA_ETI_CANTIDAD)
               rs.MoveNext
         Wend
         rs.Close
      End If
   Else
      If var_importar_codigos_master_de = 1 Then
         On Error GoTo salir
         strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb); DBQ=c:\SISTEMAS\master.xls"
         rsaux2.Open "SELECT * FROM [master$]", strConnectionString
         While Not rsaux2.EOF
               rsaux3.Open "select * from tb_articulos where vcha_art_Articulo_id = '" + IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  rs.Open "SELECT * FROM TB_ETIQUETAS_MASTER WHERE VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux2!MASTER), "", rsaux2!MASTER)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "update TB_ETIQUETAS_MASTER set floa_eti_Cantidad = " + CStr(IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)) + " where VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux2!MASTER), "", rsaux2!MASTER)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO) + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "insert into tb_etiquetas_master (vcha_eti_etiqueta_master, vcha_art_articulo_id, floa_eti_cantidad) values ('" + CStr(IIf(IsNull(rsaux2!MASTER), "", rsaux2!MASTER)) + "', '" + UCase(IIf(IsNull(rsaux2!CODIGO), "", rsaux2!CODIGO)) + "', " + CStr(IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Close
               End If
               rsaux3.Close
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         Me.lv_etiquetas.ListItems.Clear
         rs.Open "select vcha_eti_etiqueta_master, a.vcha_art_articulo_id, vcha_art_nombre_español, floa_eti_cantidad from tb_etiquetas_master a, TB_ARTICULOS b where  a.vcha_art_articulo_id = b.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_etiquetas.ListItems.Add(, , rs!vcha_eti_etiqueta_master)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id)
               list_item.SubItems(2) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ETI_CANTIDAD), "", rs!FLOA_ETI_CANTIDAD)
               rs.MoveNext
         Wend
         rs.Close
         MsgBox "Se termino la carga", vbOKOnly, "ATENCION"
      End If
      If var_importar_codigos_master_de = 2 Then
         'On Error GoTo salir2
         rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux.Open "ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
         var_cadena = "select FAC.SEGMENT1, NUMB_LMI_CANTIDAD, lec.VCHA_CAJ_CAJA_ID from XXVIA_VW_FACTURAS_CANTIA fac, XXVIA.XXVIA_TB_LEC_MOV_INV lec where fac.CUSTOMER_TRX_ID = lec.VCHA_LMI_LOCALIZADOR and fac.SEGMENT1 = lec.VCHA_LMI_CODIGO  AND TRX_DATE >= TO_DATE('01/06/2018','DD/MM/YYYY')"
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  rsaux2.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE substring(VCHA_EQU_CODIGO_EQUIVALENTE,1,8) = '" + rsaux!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     
                     If Mid(rsaux!segment1, 1, 2) = "00" Then
                        If Len(IIf(IsNull(rsaux2!vcha_art_articulo_id), "", rsaux2!vcha_art_articulo_id)) <= 8 Then
                           'VER LO DE LAS EQUIVALENCIAS
                           'var_articulo_enviado = Mid(IIf(IsNull(rsaux2!vcha_Art_Articulo_id), "", rsaux2!vcha_Art_Articulo_id), 1, 5)
                           'Else
                           var_articulo_enviado = IIf(IsNull(rsaux2!vcha_art_articulo_id), "", rsaux2!vcha_art_articulo_id)
                        End If
                     Else
                        var_articulo_enviado = IIf(IsNull(rsaux2!vcha_art_articulo_id), "", rsaux2!vcha_art_articulo_id)
                     End If
                  End If
                  rsaux2.Close
                     
                  If var_articulo_enviado <> "" Then
                     rsaux11.Open "SELECT * FROM TB_ETIQUETAS_MASTER WHERE VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux!VCHA_CAJ_CAJA_ID), "", rsaux!VCHA_CAJ_CAJA_ID)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(var_articulo_enviado), "", var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux11.EOF Then
                        rsaux9.Open "update TB_ETIQUETAS_MASTER set floa_eti_Cantidad = " + CStr(IIf(IsNull(rsaux!NUMB_LMI_CANTIDAD), 0, rsaux!NUMB_LMI_CANTIDAD)) + " where VCHA_ETI_ETIQUETA_MASTER = '" + CStr(IIf(IsNull(rsaux!VCHA_CAJ_CAJA_ID), "", rsaux!VCHA_CAJ_CAJA_ID)) + "' AND VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(var_articulo_enviado), "", var_articulo_enviado) + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux9.Open "insert into tb_etiquetas_master (vcha_eti_etiqueta_master, vcha_art_articulo_id, floa_eti_cantidad) values ('" + CStr(IIf(IsNull(rsaux!VCHA_CAJ_CAJA_ID), "", rsaux!VCHA_CAJ_CAJA_ID)) + "', '" + UCase(IIf(IsNull(var_articulo_enviado), "", var_articulo_enviado)) + "', " + CStr(IIf(IsNull(rsaux!NUMB_LMI_CANTIDAD), 0, rsaux!NUMB_LMI_CANTIDAD)) + ")", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux11.Close
                  End If
                  rsaux.MoveNext
            Wend
         End If
         rsaux.Close
         Me.lv_etiquetas.ListItems.Clear
         rs.Open "select vcha_eti_etiqueta_master, a.vcha_art_articulo_id, vcha_art_nombre_español, floa_eti_cantidad from tb_etiquetas_master a, TB_ARTICULOS b where  a.vcha_art_articulo_id = b.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_etiquetas.ListItems.Add(, , rs!vcha_eti_etiqueta_master)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id)
               list_item.SubItems(2) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ETI_CANTIDAD), "", rs!FLOA_ETI_CANTIDAD)
               rs.MoveNext
         Wend
         rs.Close
         MsgBox "Se termino de cargar los códigos master", vbOKOnly, "ATENCION"
      End If
   End If
   Exit Sub
salir2:
   MsgBox "A surgido un error al leer los códigos desde oracle", vbOKOnly, "ATENCION"
   'MsgBox Err.Description
   If rsaux2.State Then
      rsaux2.Close
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If

Exit Sub
salir:
   MsgBox "A surgido un error al leer el archivo puede que este no tena la estructura correcta que es: Nombre del archivo MASTER, Nombre de la hoja MASTER, Nombre de las columnas MASTER, CODIGO, CANTIDAD", vbOKOnly, "ATENCION"
   If rsaux2.State Then
      rsaux2.Close
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   Exit Sub
End Sub

Private Sub cmd_eliminar_Click()
   var_si = MsgBox("¿Desea eliminar la etiqueta master?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la eliminación de la etiqueta master", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rsaux.Open "delete from tb_etiquetas_master where vcha_eti_etiqueta_master = '" + Me.lv_etiquetas.selectedItem + "' and vcha_art_articulo_id = '" + Me.lv_etiquetas.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
         lv_etiquetas.ListItems.Remove (lv_etiquetas.selectedItem.Index)
      End If
   End If
   
End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_master <> "" Then
      If Me.txt_codigo <> "" Then
         If IsNumeric(Me.txt_cantidad) Then
            rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "SELECT * FROM TB_ETIQUETAS_MASTER WHERE VCHA_ETI_ETIQUETA_MASTER = '" + Me.txt_master + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  rsaux1.Open "INSERT INTO TB_ETIQUETAS_MASTER (VCHA_ETI_ETIQUETA_MASTER, VCHA_ART_ARTICULO_ID, FLOA_ETI_CANTIDAD) VALUES ('" + Me.txt_master + "','" + Me.txt_codigo + "'," + CStr(CDbl(Me.txt_cantidad)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = lv_etiquetas.ListItems.Add(, , Me.txt_master)
                  list_item.SubItems(1) = Me.txt_codigo
                  list_item.SubItems(2) = Me.txt_descripcion
                  list_item.SubItems(3) = Me.txt_cantidad
               Else
                  var_si = MsgBox("¿Se actualizara la cantidad de la etiqueta master?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     rsaux2.Open "UPDATE TB_ETIQUETAS_MASTER SET FLOA_ETI_CANTIDAD = " + Me.txt_cantidad + " WHERE VCHA_ETI_ETIQUETA_MASTER = '" + Me.txt_master + "' AND VCHA_ART_aRTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     Me.lv_etiquetas.selectedItem = CStr(Me.txt_master)
                     Me.lv_etiquetas.selectedItem.SubItems(1) = Me.txt_codigo
                     Me.lv_etiquetas.selectedItem.SubItems(2) = Me.txt_descripcion
                     Me.lv_etiquetas.selectedItem.SubItems(3) = Me.txt_cantidad
                  End If
                  rsaux1.Close
               End If
               rsaux.Close
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Código master incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set a = fs.CreateTextFile(App.Path + "\etiqueta.txt", True)
   a.writeline ("")
   a.writeline ("US")
   a.writeline ("N")
   a.writeline ("q816")
   a.writeline ("Q1015,20+0")
   a.writeline ("S2")
   a.writeline ("D8")
   a.writeline ("ZT")
   a.writeline ("TTh:m")
   a.writeline ("TDy2.mn.dd")
   a.writeline ("A605,20,1,4,2,1,N,""" + Mid(Me.lv_etiquetas.selectedItem.SubItems(2), 1, 47) + """")
   a.writeline ("A505,20,1,4,2,2,N,""" + "SKU: " + Me.lv_etiquetas.selectedItem.SubItems(1) + """")
   a.writeline ("A400,20,1,4,2,2,N,""" + "CANTIDAD: " + Me.lv_etiquetas.selectedItem.SubItems(3) + """")
   a.writeline ("B77,782,0,1,4,9,101,B,""" + Me.lv_etiquetas.selectedItem + """")
   a.writeline ("P1")
   a.Close
   Open (App.Path & "\etiquetas.bat") For Output As #2
   var_Archivo = App.Path & "\etiquetas.bat"
   Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
   Close #2
   x = Shell(var_Archivo, vbHide)
   
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_cantidad = ""
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_master = ""
   Me.txt_master.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
            Set reporte = appl.OpenReport(App.Path + "\rep_codigos_master.rpt")
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de ventas"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_codigos_master.rpt")
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Codigos_master_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
End Sub

Private Sub Form_Load()
   Top = 200
   Left = 1000
   rs.Open "select vcha_eti_etiqueta_master, a.vcha_art_articulo_id, vcha_art_nombre_español, floa_eti_cantidad from tb_etiquetas_master a, TB_ARTICULOS b where  a.vcha_art_articulo_id = b.VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_etiquetas.ListItems.Add(, , rs!vcha_eti_etiqueta_master)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id)
         list_item.SubItems(2) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
         list_item.SubItems(3) = IIf(IsNull(rs!FLOA_ETI_CANTIDAD), "", rs!FLOA_ETI_CANTIDAD)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_acumulado_ventas)
End Sub

Private Sub lv_etiquetas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_etiquetas, ColumnHeader)
End Sub

Private Sub lv_etiquetas_ItemClick(ByVal item As MSComctlLib.ListItem)
   Me.txt_master = Me.lv_etiquetas.selectedItem
   Me.txt_codigo = Me.lv_etiquetas.selectedItem.SubItems(1)
   Me.txt_descripcion = Me.lv_etiquetas.selectedItem.SubItems(2)
   Me.txt_cantidad = Me.lv_etiquetas.selectedItem.SubItems(3)
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
         Me.txt_descripcion.SetFocus
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_master_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
