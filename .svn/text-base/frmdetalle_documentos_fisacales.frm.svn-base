VERSION 5.00
Begin VB.Form frmdetalle_documentos_fiscales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Documentos Fiscales"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7005
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   0
      TabIndex        =   11
      Top             =   330
      Width           =   6975
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6480
      Picture         =   "frmdetalle_documentos_fisacales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir Esc"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento a cancelar "
      Height          =   1455
      Left            =   90
      TabIndex        =   1
      Top             =   465
      Width           =   6825
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   645
         Width           =   795
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   990
         Width           =   1035
      End
      Begin VB.ComboBox cmb_documentos 
         Height          =   315
         ItemData        =   "frmdetalle_documentos_fisacales.frx":063A
         Left            =   2550
         List            =   "frmdetalle_documentos_fisacales.frx":0644
         TabIndex        =   3
         Top             =   300
         Width           =   4155
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   495
         TabIndex        =   9
         Top             =   705
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   495
         TabIndex        =   8
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento: "
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lbl_estatus 
         Caption         =   "Label3"
         Height          =   210
         Left            =   2745
         TabIndex        =   6
         Top             =   1050
         Width           =   3465
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmdetalle_documentos_fisacales.frx":0662
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Refacturar"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmdetalle_documentos_fiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_documentos_Click()
   txt_numero = ""
   var_estatus = ""
   lbl_estatus = ""
   If cmb_documentos = "FACTURA" Then
      txt_documento = "FA"
   End If
   If cmb_documentos = "NOTA DE CREDITO" Then
      txt_documento = "NC"
   End If
End Sub

Private Sub cmb_documentos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_series.Enabled = True Then
         cmb_series.SetFocus
      Else
         txt_numero.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_series_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_numero.SetFocus
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 2200
   lbl_estatus = ""
   var_cadena_seguridad = ""
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_documento.Enabled = True
      cmb_documentos.Enabled = True
      txt_numero.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!vcha_ser_serie_id
      var_serie = rs!vcha_ser_serie_id
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_documento.Enabled = False
      cmb_documentos.Enabled = False
      txt_numero.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_detalle_documentos_fiscales)
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      cmb_documentos.SetFocus
   End If
End Sub

Private Sub txt_documento_LostFocus()
   lbl_estatus = ""
   txt_numero = ""
   var_tipo_facturacion = ""
   If Trim(txt_documento) <> "" Then
      If txt_documento = "FA" Then
         cmb_documentos = "FACTURA"
         var_tipo_facturacion = ""
      Else
         If txt_documento = "NC" Then
            cmb_documentos = "NOTA DE CREDITO"
         Else
            MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
            txt_documento = ""
            cmb_documentos = ""
         End If
      End If
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Trim(txt_documento) <> "" Then
      If Not IsNumeric(txt_numero) Then
         MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
         txt_numero = ""
      Else
         var_estatus = ""
         rs.Open "select isnull(char_car_estatus,'') as char_car_estatus, isnull(char_car_tipo_facturacion,'') as char_car_tipo_facturacion, isnull(vcha_car_documento,'') as vcha_car_documento from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + txt_documento + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!CHAR_CAR_ESTATUS = "" Or rs!CHAR_CAR_ESTATUS = " " Then
               lbl_estatus = "ESTATUS: IMPRESA"
               var_estatus = "I"
               If txt_documento = "FA" Then
                  var_tipo_facturacion = rs!char_Car_tipo_facturacion
               End If
               If txt_documento = "NC" Then
                  var_tipo_nota_credito = rs!vcha_Car_documento
               End If
            End If
            If rs!CHAR_CAR_ESTATUS = "C" Then
               lbl_estatus = "ESTATUS: CANCELADA"
               var_estatus = "C"
               If txt_documento = "FA" Then
                  var_tipo_facturacion = rs!char_Car_tipo_facturacion
               End If
               If txt_documento = "NC" Then
                  var_tipo_nota_credito = rs!vcha_Car_documento
               End If
            End If
         Else
            lbl_estatus = "ESTATUS: NO IMPRESA"
            var_estatus = "N"
         End If
         rs.Close
      End If
   Else
      MsgBox "Se debe de seleccionar un tipo de documento", vbOKOnly, "ATENCION"
      txt_numero = ""
   End If
End Sub
