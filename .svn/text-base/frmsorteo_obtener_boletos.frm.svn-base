VERSION 5.00
Begin VB.Form frmsorteo_obtener_boletos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de boletos"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmsorteo_obtener_boletos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Aceptar cambios Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6405
      Picture         =   "frmsorteo_obtener_boletos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmsorteo_obtener_boletos.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualizar boletos "
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   15
      TabIndex        =   4
      Top             =   315
      Width           =   6765
   End
   Begin VB.Frame Frame1 
      Caption         =   " Folios "
      Height          =   870
      Left            =   75
      TabIndex        =   3
      Top             =   435
      Width           =   6645
      Begin VB.TextBox txt_actual 
         Height          =   350
         Left            =   5130
         TabIndex        =   2
         Top             =   327
         Width           =   1410
      End
      Begin VB.TextBox txt_al 
         Height          =   350
         Left            =   2805
         TabIndex        =   1
         Top             =   327
         Width           =   1410
      End
      Begin VB.TextBox txt_del 
         Height          =   350
         Left            =   570
         TabIndex        =   0
         Top             =   327
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Actual:"
         Height          =   195
         Left            =   4620
         TabIndex        =   9
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   2550
         TabIndex        =   8
         Top             =   405
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   405
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmsorteo_obtener_boletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   rs.Open "select NUMB_ABO_INICIO_BOLETOS , NUMB_ABO_TERMINO_BOLETOS  from tb_asignacion_boletos where vcha_esb_establecimiento_id = 'E000001666'", cnnsorteo, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_del = IIf(IsNull(rs!NUMB_ABO_INICIO_BOLETOS), "", rs!NUMB_ABO_INICIO_BOLETOS)
      Me.txt_al = IIf(IsNull(rs!NUMB_ABO_TERMINO_BOLETOS), "", rs!NUMB_ABO_TERMINO_BOLETOS)
      Me.txt_actual = IIf(IsNull(rs!NUMB_ABO_INICIO_BOLETOS), "", rs!NUMB_ABO_INICIO_BOLETOS)
   Else
      MsgBox "No existen folios asignados al almacen", vbOKOnly, "ATENCION"
   End If
   rs.Close
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If IsNumeric(Me.txt_del) Then
      If IsNumeric(Me.txt_al) Then
         rs.Open "select * from tb_sorteo_folios", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "update tb_sorteo_folios SET INTE_SOR_FOLIO_INICIO = " + Me.txt_del + ", INTE_SOR_FOLIO_FIN = " + Me.txt_al + ", INTE_SOR_FOLIO_ACTUAL = " + Me.txt_actual, cnn, adOpenDynamic, adLockOptimistic
         Else
            rsaux.Open "INSERT INTO TB_SORTEO_FOLIOS (INTE_SOR_FOLIO_INICIO, INTE_SOR_FOLIO_FIN, INTE_SOR_FOLIO_ACTUAL) VALUES ( " + Me.txt_del + "," + Me.txt_al + ", " + Me.txt_actual + ")", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Close
         MsgBox "Se han actualizado los boletos", vbOKOnly, "ATENCION"
      Else
         MsgBox "Numero de folio final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de folio inicial incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 2800
   If cnnsorteo.State = 0 Then
      cnnsorteo.Open var_conexion_sorteo
      cnnsorteo.CursorLocation = adUseClient
   End If
   rs.Open "select * from tb_sorteo_folios", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_del = IIf(IsNull(rs!inte_sor_folio_inicio), "", rs!inte_sor_folio_inicio)
      Me.txt_al = IIf(IsNull(rs!inte_sor_folio_fin), "", rs!inte_sor_folio_fin)
      Me.txt_actual = IIf(IsNull(rs!inte_sor_folio_actual), "", rs!inte_sor_folio_actual)
   End If
   rs.Close
   'rs.Open "select NUMB_ABO_INICIO_BOLETOS , NUMB_ABO_TERMINO_BOLETOS  from tb_asignacion_boletos where vcha_esb_establecimiento_id = 'E000001666'", cnnsorteo, adOpenDynamic, adLockOptimistic
   'If Not rs.EOF Then
   '   Me.txt_del = IIf(IsNull(rs!NUMB_ABO_INICIO_BOLETOS), "", rs!NUMB_ABO_INICIO_BOLETOS)
   '   Me.txt_al = IIf(IsNull(rs!NUMB_ABO_TERMINO_BOLETOS), "", rs!NUMB_ABO_TERMINO_BOLETOS)
   'End If
   'rs.Close
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_al_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_del_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub
