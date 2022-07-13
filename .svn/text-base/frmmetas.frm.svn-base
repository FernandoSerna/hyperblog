VERSION 5.00
Begin VB.Form frmmetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Metas"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4815
      Picture         =   "frmmetas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmmetas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmmetas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Guardar Alt + G"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmmetas.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmmetas.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   105
      TabIndex        =   10
      Top             =   315
      Width           =   5025
   End
   Begin VB.Frame Frame1 
      Caption         =   " Meta "
      Height          =   1170
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   5025
      Begin VB.TextBox txt_meta 
         Height          =   330
         Left            =   3495
         TabIndex        =   9
         Top             =   675
         Width           =   1395
      End
      Begin VB.TextBox txt_mes 
         Height          =   330
         Left            =   2010
         TabIndex        =   7
         Top             =   675
         Width           =   840
      End
      Begin VB.TextBox txt_año 
         Height          =   330
         Left            =   630
         TabIndex        =   5
         Top             =   675
         Width           =   825
      End
      Begin VB.TextBox txt_nombre_ruta 
         Height          =   330
         Left            =   1470
         TabIndex        =   3
         Top             =   300
         Width           =   3420
      End
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   645
         TabIndex        =   2
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Meta:"
         Height          =   195
         Left            =   2955
         TabIndex        =   8
         Top             =   743
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Left            =   1620
         TabIndex        =   6
         Top             =   743
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   743
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   368
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmmetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_eliminar_Click()
   If Me.txt_ruta <> "" Then
      If IsNumeric(Me.txt_año) Then
         If IsNumeric(Me.txt_mes) Then
            var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rs.Open "DELETE FROM TB_METAS_COMISIONES WHERE VCHA_RUT_RUTA_ID  ='" + Me.txt_ruta + "' AND INTE_MET_AÑO = " + Me.txt_año + " AND INTE_MET_MES = " + Me.txt_mes, cnn, adOpenDynamic, adLockOptimistic
               MsgBox "Se a eliminado el registro", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Año incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado una ruta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_ruta <> "" Then
      If IsNumeric(Me.txt_año) Then
         If IsNumeric(Me.txt_mes) Then
            If CDbl(Me.txt_mes) <= 12 Then
               If IsNumeric(Me.txt_meta) Then
                  rs.Open "SELECT * FROM TB_METAS_COMISIONES WHERE VCHA_RUT_RUTA_ID = '" + Me.txt_ruta + "' AND inte_met_año = " + Me.txt_año + " and inte_met_mes = " + Me.txt_mes, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     rsaux.Open "update TB_METAS_COMISIONES set floa_met_meta = " + Me.txt_meta + " WHERE VCHA_RUT_RUTA_ID = '" + Me.txt_ruta + "' AND inte_met_año = " + Me.txt_año + " and inte_met_mes = " + Me.txt_mes, cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se han aplicado los cambios", vbOKOnly, "ATENCION"
                  Else
                     rsaux.Open "insert into TB_METAS_COMISIONES (VCHA_RUT_RUTA_ID, INTE_MET_AÑO, INTE_MET_MES, FLOA_MET_META) values ('" + Me.txt_ruta + "', " + Me.txt_año + ", " + Me.txt_mes + "," + Me.txt_meta + ")", cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
                  Me.txt_mes.SetFocus
                  
               Else
                  MsgBox "Meta incorrecta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Año incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Debe de indicar una ruta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_ruta = ""
   Me.txt_año = ""
   Me.txt_mes = ""
   Me.txt_meta = ""
   Me.txt_nombre_ruta = ""
   Me.txt_ruta.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_año_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_mes_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_mes_LostFocus()
   Me.txt_meta = ""
End Sub

Private Sub txt_meta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_nombre_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_LostFocus()
   If Me.txt_ruta <> "" Then
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + Me.txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_ruta = IIf(IsNull(rs!VCHA_RUT_NOMBRE), "", rs!VCHA_RUT_NOMBRE)
      Else
         MsgBox "Ruta incorrecta", vbOKOnly, "ATENCIO"
         Me.txt_nombre_ruta = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_ruta = ""
   End If
End Sub
