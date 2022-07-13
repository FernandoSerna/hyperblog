VERSION 5.00
Begin VB.Form frmlista_precios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Precios"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmlista_precios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5130
      Picture         =   "frmlista_precios.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmlista_precios.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Artículo "
      Height          =   1665
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   5370
      Begin VB.TextBox txt_nombre_lista_precios 
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   210
         Width           =   3750
      End
      Begin VB.TextBox txt_lista_precios 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   210
         Width           =   450
      End
      Begin VB.TextBox txt_precio 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1215
         Width           =   1530
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   885
         Width           =   4215
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   555
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lista:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   945
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   615
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   5400
   End
End
Attribute VB_Name = "frmlista_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_guardar_Click()
   Dim var_coddigo_auxiliar As String
   Dim var_codigo As String
   Dim verificado As Integer
   Dim var_preico As Double
   Dim VERIFICADOR As Integer
   If Me.txt_lista_precios <> "" Then
      If IsNumeric(Me.txt_precio) Then
         If var_empresa = "18" Then
            If Mid(Me.txt_codigo, 11, 1) = "0" Then
               var_codigo_auxiliar = Left(Me.txt_codigo, 10)
               For var_i = 0 To 9
                  var_codigo = Trim(var_codigo_auxiliar) + Trim(CStr(var_i))
                  sum1 = 0
                  sum2 = 0
                  mcodigo = var_codigo
                  longitud = Len(mcodigo)
                  For icont = 1 To longitud
                      If ((icont / 2) - Int((icont / 2))) = 0 Then
                         sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                      Else
                         sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                      End If
                  Next icont
                  msuma = sum1 * 13 + sum2
                  VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                  If VERIFICADOR = 10 Then
                     VERIFICADOR = 0
                  End If
                  var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                  var_precio = CDbl(Me.txt_precio)
                  VAR_dESCUENTO_ARTICULO = (100 - (CDbl(var_i) * 10)) / 100
                  var_precio = var_precio * VAR_dESCUENTO_ARTICULO
              
                  rs.Open "select * from tb_Articulos where vcha_art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     If Me.txt_lista_precios = "01" Then
                        rsaux.Open "update tb_Articulos set mone_art_precio_base = " + CStr(var_precio) + " where vcha_Art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + Me.txt_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        rsaux2.Open "update tb_detalle_lista_precios set floa_dli_precio = " + CStr(var_precio) + " where vcha_lis_lista_precios_id = '" + Me.txt_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux2.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('" + Me.txt_lista_precios + "','" + var_codigo + "'," + CStr(var_precio) + " )", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                  End If
                  rs.Close
               Next var_i
            End If
         Else
            If Trim(Me.txt_codigo) <> "" Then
               var_codigo = Me.txt_codigo
               var_precio = CDbl(Me.txt_precio)
               rs.Open "select * from tb_Articulos where vcha_art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If Me.txt_lista_precios = "01" Then
                     rsaux.Open "update tb_Articulos set mone_art_precio_base = " + CStr(var_precio) + " where vcha_Art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Open "select * from tb_detalle_lista_precios where vcha_lis_lista_precios_id = '" + Me.txt_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux2.Open "update tb_detalle_lista_precios set floa_dli_precio = " + CStr(var_precio) + " where vcha_lis_lista_precios_id = '" + Me.txt_lista_precios + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a actualizado el precio", vbOKOnly, "ATENCION"
                  Else
                     rsaux2.Open "INSERT INTO TB_DETALLE_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_DLI_PRECIO) VALUES ('" + Me.txt_lista_precios + "','" + var_codigo + "'," + CStr(var_precio) + " )", cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a insertado el precio", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
                  
               End If
               rs.Close
            End If
         End If
      Else
         MsgBox "Precio incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Debe de indicar una lista de precios", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   If Me.txt_lista_precios <> "" Then
      frmreporte_lista_precios_01.txt_lista = Me.txt_lista_precios
      frmreporte_lista_precios_01.Show 1
   Else
      MsgBox "Debe de seleccionar una lista de precios", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   If var_empresa = "18" Then
      If Len(Trim(Me.txt_codigo)) > 10 Then
         Me.txt_codigo = Mid(Me.txt_codigo, 1, 10)
      End If
      If Len(Trim(Me.txt_codigo)) = 10 Then
         rs.Open "select * from tb_articulos where vcha_Art_articulo_id like '" + Trim(Me.txt_codigo) + "0%'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_codigo = IIf(IsNull(rs!vcha_art_articulo_id), "", rs!vcha_art_articulo_id)
         End If
         rs.Close
      End If
      If Me.txt_lista_precios <> "" Then
         If Trim(Me.txt_codigo) <> "" Then
            If Mid(Me.txt_codigo, 11, 1) = "0" Then
               rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux.Open "Select * from tb_Detalle_lista_precios where VCHA_LIS_LISTA_PRECIOS_ID = '" + Me.txt_lista_precios + "' and vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_Español), "", rs!vcha_Art_nombre_Español)
                     Me.txt_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio)
                  Else
                     Me.txt_precio = IIf(IsNull(rs!mone_art_precio_base), 0, rs!mone_art_precio_base)
                     Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_Español), "", rs!vcha_Art_nombre_Español)
                     MsgBox "El artículo no se encuentra en la lista de precios", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "El artículo debe de tener descuento 0", vbOKOnly, "ATENCION"
            End If
         End If
      Else
         MsgBox "Debe de seleccionar una lista de precios", vbOKOnly, "ATENCION"
      End If
   Else
      rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_Español), "", rs!vcha_Art_nombre_Español)
         rsaux.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '" + Me.txt_lista_precios + "' and vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            Me.txt_precio = IIf(IsNull(rs!mone_art_precio_base), 0, rs!mone_art_precio_base)
         Else
            Me.txt_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio)
         End If
         rsaux.Close
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_codigo = ""
         Me.txt_descripcion = ""
         Me.txt_precio = ""
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

Private Sub txt_lista_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_lista_precios_LostFocus()
   If Me.txt_lista_precios <> "" Then
      rs.Open "SELECT * FROM TB_LISTADEPRECIOS WHERE VCHA_LIS_LISTA_ID = '" + Me.txt_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_lista_precios = IIf(IsNull(rs!VCHA_lis_NOMBRE), "", rs!VCHA_lis_NOMBRE)
      Else
         rsaux.Open "select * from TB_LISTADEPRECIOS", cnn, adOpenDynamic, adLockOptimistic
         var_cadena = ""
         While Not rsaux.EOF
               If var_cadena = "" Then
                  var_cadena = var_cadena + IIf(IsNull(rsaux!vcha_lis_lista_id), "", rsaux!vcha_lis_lista_id) + " " + IIf(IsNull(rsaux!VCHA_lis_NOMBRE), "", rsaux!VCHA_lis_NOMBRE)
               Else
                  var_cadena = var_cadena + ", " + IIf(IsNull(rsaux!vcha_lis_lista_id), "", rsaux!vcha_lis_lista_id) + " " + IIf(IsNull(rsaux!VCHA_lis_NOMBRE), "", rsaux!VCHA_lis_NOMBRE)
               End If
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.txt_lista_precios = ""
         Me.txt_nombre_lista_precios = ""
         MsgBox "La lista de precios no existe, debe de seleccionar " + var_cadena, vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_lista_precios = ""
   End If
End Sub

Private Sub txt_nombre_lista_precios_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_lista_precios.SetFocus
   End If
End Sub

Private Sub txt_nombre_lista_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
  ' MsgBox CStr(KeyAscii)
  
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 45
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
