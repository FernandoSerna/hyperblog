VERSION 5.00
Begin VB.Form frmimpresion_etiqueta_coppel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de etiquetas de COPPEL"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6270
   Begin VB.Frame Frame3 
      Caption         =   "Datos de la etiqueta "
      Height          =   3885
      Left            =   135
      TabIndex        =   7
      Top             =   2460
      Width           =   5985
      Begin VB.TextBox txt_et_dato10 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   19
         Top             =   3240
         Width           =   1635
      End
      Begin VB.TextBox txt_et_dato12 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2730
         TabIndex        =   18
         Top             =   3315
         Width           =   2370
      End
      Begin VB.TextBox txt_et_dato11 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2730
         TabIndex        =   17
         Top             =   2820
         Width           =   2040
      End
      Begin VB.TextBox txt_et_dato9 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   16
         Top             =   2755
         Width           =   1290
      End
      Begin VB.TextBox txt_et_dato8 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2730
         TabIndex        =   15
         Top             =   2325
         Width           =   1995
      End
      Begin VB.TextBox txt_et_dato7 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   14
         Top             =   2270
         Width           =   1710
      End
      Begin VB.TextBox txt_et_dato6 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2730
         TabIndex        =   13
         Top             =   1785
         Width           =   1995
      End
      Begin VB.TextBox txt_et_dato5 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   12
         Top             =   1785
         Width           =   1680
      End
      Begin VB.TextBox txt_et_dato4 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2730
         TabIndex        =   11
         Top             =   1290
         Width           =   1860
      End
      Begin VB.TextBox txt_et_dato3 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   10
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox txt_et_dato2 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   9
         Top             =   815
         Width           =   3090
      End
      Begin VB.TextBox txt_et_dato1 
         Enabled         =   0   'False
         Height          =   350
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Factura "
      Height          =   1050
      Left            =   135
      TabIndex        =   2
      Top             =   1305
      Width           =   5985
      Begin VB.TextBox txt_piezas 
         Enabled         =   0   'False
         Height          =   350
         Left            =   3885
         TabIndex        =   6
         Top             =   387
         Width           =   1200
      End
      Begin VB.TextBox txt_factura 
         Enabled         =   0   'False
         Height          =   350
         Left            =   1470
         TabIndex        =   4
         Top             =   387
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         Height          =   195
         Left            =   3360
         TabIndex        =   5
         Top             =   465
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   795
         TabIndex        =   3
         Top             =   465
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Caja "
      Height          =   1140
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5985
      Begin VB.TextBox txt_caja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   675
         Left            =   1380
         TabIndex        =   1
         Top             =   315
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmimpresion_etiqueta_coppel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text10_Change()

End Sub

Private Sub Form_Load()
   Top = 500
   Left = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   On Error GoTo salir:
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      VAR_EMBARQUE = CDbl(Mid(Me.txt_caja, 2, 6))
      var_caja = CDbl(Mid(Me.txt_caja, 8, 3))
      cnn.CommandTimeout = 360
      rs.Open "select * from tb_detalle_cajas with (nolock)  where inte_emb_embarque = " + CStr(VAR_EMBARQUE) + " and inte_paq_caja = " + CStr(var_caja), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
         rsaux.Open "SELECT * FROM TB_ENC_ORDEN_SURTIDO WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_pedido = rsaux!inte_ped_numero
            rsaux2.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               VAR_PEDIDO_COPPEL = IIf(IsNull(rsaux2!VCHA_PED_PEDIDO_EXTERNO), "", rsaux2!VCHA_PED_PEDIDO_EXTERNO)
               If Trim(VAR_PEDIDO_COPPEL) <> "" Then
                  var_establecimiento = IIf(IsNull(rsaux2!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux2!vcha_ESB_ESTABLECIMIENTO_id)
                  rsaux3.Open "select * from tb_establecimientos where vcha_esb_Establecimiento_id = '" + var_establecimiento + "' ", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_establecimiento_coppel = IIf(IsNull(rsaux3!vcha_esb_establecimiento_anterior_id), "", rsaux3!vcha_esb_establecimiento_anterior_id)
                     If Trim(var_establecimiento_coppel) <> "" Then
                        rsaux4.Open "select * from TB_PEDIDO_ORIGINAL_COPPEL where archivo = '" + VAR_PEDIDO_COPPEL + "' and destino = '" + var_establecimiento_coppel + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           VAR_PEDIDO_ORIGINAL = rsaux4!NUMPEDIDO
                           Me.txt_et_dato1 = rsaux4!ET_DATO1
                           Me.txt_et_dato2 = rsaux4!ET_DATO2
                           Me.txt_et_dato3 = rsaux4!ET_DATO3
                           Me.txt_et_dato4 = rsaux4!ET_DATO4
                           Me.txt_et_dato5 = rsaux4!ET_DATO5
                           Me.txt_et_dato6 = rsaux4!ET_DATO6
                           Me.txt_et_dato7 = rsaux4!ET_DATO7
                           Me.txt_et_dato8 = rsaux4!ET_DATO8
                           Me.txt_et_dato9 = rsaux4!ET_DATO9
                           Me.txt_et_dato10 = rsaux4!ET_DATO10
                           Me.txt_et_dato11 = rsaux4!ET_DATO11
                           Me.txt_et_dato12 = rsaux4!ET_DATO12
                           rsaux5.Open "select * from VW_ETIQUETA_COPPEL where INTE_EMO_NUMERO_ORIGEN = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              Me.txt_et_dato6 = "FACTURA: " + CStr(rsaux5!inte_Car_numero)
                              Me.txt_et_dato7 = "UNIDS.FACTURA: " + CStr(rsaux5!Cantidad)
                              Me.txt_factura = CStr(rsaux5!inte_Car_numero)
                              Me.txt_piezas = rsaux5!Cantidad
                           End If
                           rsaux5.Close
                           rsaux5.Open "select max(inte_paq_caja) from tb_Detalle_cajas where inte_emb_embarque = " + CStr(VAR_EMBARQUE) + " and inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              var_maximo_numero = rsaux5(0).Value
                           End If
                           rsaux5.Close
                           var_caja_STR = CStr(var_caja)
                           VAR_MAXIMO_cAJA = CStr(var_maximo_numero)
                           If Len(var_caja_STR) = 1 Then
                              var_caja_STR = "00" + var_caja_STR
                           Else
                              If Len(var_caja_STR) = 2 Then
                                 var_caja_STR = "0" + var_caja_STR
                              End If
                           End If
                           If Len(VAR_MAXIMO_cAJA) = 1 Then
                              VAR_MAXIMO_cAJA = "00" + VAR_MAXIMO_cAJA
                           Else
                              If Len(VAR_MAXIMO_cAJA) = 2 Then
                                 VAR_MAXIMO_cAJA = "0" + VAR_MAXIMO_cAJA
                              End If
                           End If
                           Me.txt_et_dato10 = Trim(var_caja_STR) + "/" + Trim(VAR_MAXIMO_cAJA)
                           If Len(var_establecimiento_coppel) = 1 Then
                              var_establecimiento_coppel = "0" + var_establecimiento_coppel
                           End If
                           Me.txt_et_dato11 = "P" + Trim(VAR_PEDIDO_ORIGINAL) + "A" + Trim(var_establecimiento_coppel) + Trim(var_caja_STR) + Trim(VAR_MAXIMO_cAJA)
                           Me.txt_et_dato12 = "P " + Trim(VAR_PEDIDO_ORIGINAL) + "A" + " " + Trim(var_establecimiento_coppel) + " " + Trim(var_caja_STR) + " " + Trim(VAR_MAXIMO_cAJA)
                           Set fs = CreateObject("Scripting.FileSystemObject")
                           Set a = fs.CreateTextFile(App.Path + "\etiquetas.txt", True)
a.writeline ("")
a.writeline ("US")
a.writeline ("N")
a.writeline ("q816")
a.writeline ("Q608,24")
a.writeline ("S2")
a.writeline ("D10")
a.writeline ("ZT")
a.writeline ("A30,22,0,4,2,2,N," + """" + Me.txt_et_dato1 + """")
a.writeline ("A30,87,0,5,1,1,N," + """" + Me.txt_et_dato2 + """")
a.writeline ("A30,154,0,4,1,2,N," + """" + Me.txt_et_dato3 + """")
a.writeline ("A564,154,0,4,1,2,N," + """" + Me.txt_et_dato4 + """")
a.writeline ("A30,211,0,3,1,2,N," + """" + Me.txt_et_dato5 + """")
a.writeline ("A30,258,0,3,1,2,N," + """" + Me.txt_et_dato7 + """")
a.writeline ("A371,211,0,3,1,2,N," + """" + Me.txt_et_dato6 + """")
a.writeline ("A367,258,0,3,1,2,N," + """" + Me.txt_et_dato8 + """")
a.writeline ("LO6,305,811,4")
a.writeline ("LO298,305,5,282")
a.writeline ("A108,324,0,3,1,2,N," + """" + Me.txt_et_dato9 + """")
a.writeline ("A32,391,0,5,1,3,N," + """" + Me.txt_et_dato10 + """")
a.writeline ("B371,317,0,1A,2,2,185,N," + """" + Me.txt_et_dato11 + """")
a.writeline ("A412,514,0,3,1,2,N," + """" + Me.txt_et_dato12 + """")
a.writeline ("P1")
                           
                           
                          
                           a.Close
                           Open (App.Path & "\etiquetas.bat") For Output As #2
                           var_Archivo = App.Path & "\etiquetas.bat"
                           Print #2, "copy " + App.Path + "\etiquetas.txt lpt1"
                           Close #2
                           x = Shell(var_Archivo, vbHide)
                           Me.txt_caja = ""
                        Else
                           MsgBox "Establecimiento incorrecto", vbOKOnly, "ATENCION"
                           Me.txt_et_dato1 = ""
                           Me.txt_et_dato2 = ""
                           Me.txt_et_dato3 = ""
                           Me.txt_et_dato4 = ""
                           Me.txt_et_dato5 = ""
                           Me.txt_et_dato6 = ""
                           Me.txt_et_dato7 = ""
                           Me.txt_et_dato8 = ""
                           Me.txt_et_dato9 = ""
                           Me.txt_et_dato10 = ""
                           Me.txt_et_dato11 = ""
                           Me.txt_et_dato12 = ""
                        End If
                        rsaux4.Close
                     Else
                        MsgBox "Establecimiento incorrecto", vbOKOnly, "ATENCION"
                        Me.txt_et_dato1 = ""
                        Me.txt_et_dato2 = ""
                        Me.txt_et_dato3 = ""
                        Me.txt_et_dato4 = ""
                        Me.txt_et_dato5 = ""
                        Me.txt_et_dato6 = ""
                        Me.txt_et_dato7 = ""
                        Me.txt_et_dato8 = ""
                        Me.txt_et_dato9 = ""
                        Me.txt_et_dato10 = ""
                        Me.txt_et_dato11 = ""
                        Me.txt_et_dato12 = ""
                     End If
                  End If
                  rsaux3.Close
               Else
                  MsgBox "Pedido invalido", vbOKOnly, "ATENCION"
                  Me.txt_et_dato1 = ""
                  Me.txt_et_dato2 = ""
                  Me.txt_et_dato3 = ""
                  Me.txt_et_dato4 = ""
                  Me.txt_et_dato5 = ""
                  Me.txt_et_dato6 = ""
                  Me.txt_et_dato7 = ""
                  Me.txt_et_dato8 = ""
                  Me.txt_et_dato9 = ""
                  Me.txt_et_dato10 = ""
                  Me.txt_et_dato11 = ""
                  Me.txt_et_dato12 = ""
               End If
            Else
               MsgBox "Pedido invalido", vbOKOnly, "ATENCION"
               Me.txt_et_dato1 = ""
               Me.txt_et_dato2 = ""
               Me.txt_et_dato3 = ""
               Me.txt_et_dato4 = ""
               Me.txt_et_dato5 = ""
               Me.txt_et_dato6 = ""
               Me.txt_et_dato7 = ""
               Me.txt_et_dato8 = ""
               Me.txt_et_dato9 = ""
               Me.txt_et_dato10 = ""
               Me.txt_et_dato11 = ""
               Me.txt_et_dato12 = ""
            End If
            rsaux2.Close
         Else
            MsgBox "Orden de surtido invalida", vbOKOnly, "ATENCION"
            Me.txt_et_dato1 = ""
            Me.txt_et_dato2 = ""
            Me.txt_et_dato3 = ""
            Me.txt_et_dato4 = ""
            Me.txt_et_dato5 = ""
            Me.txt_et_dato6 = ""
            Me.txt_et_dato7 = ""
            Me.txt_et_dato8 = ""
            Me.txt_et_dato9 = ""
            Me.txt_et_dato10 = ""
            Me.txt_et_dato11 = ""
            Me.txt_et_dato12 = ""
         End If
         rsaux.Close
      Else
         MsgBox "Caja invalida", vbOKOnly, "ATENCION"
         Me.txt_et_dato1 = ""
         Me.txt_et_dato2 = ""
         Me.txt_et_dato3 = ""
         Me.txt_et_dato4 = ""
         Me.txt_et_dato5 = ""
         Me.txt_et_dato6 = ""
         Me.txt_et_dato7 = ""
         Me.txt_et_dato8 = ""
         Me.txt_et_dato9 = ""
         Me.txt_et_dato10 = ""
         Me.txt_et_dato11 = ""
         Me.txt_et_dato12 = ""
      End If
      rs.Close
   End If
Exit Sub
salir:
  If rs.State = 1 Then
     rs.Close
  End If
  If rsaux.State = 1 Then
     rsaux.Close
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
End Sub
