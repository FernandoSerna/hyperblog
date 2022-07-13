VERSION 5.00
Begin VB.Form frmetiquetas_ubicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmetiquetas_ubicaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6975
      Picture         =   "frmetiquetas_ubicaciones.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   105
      TabIndex        =   10
      Top             =   285
      Width           =   7230
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del artículo "
      Height          =   2010
      Left            =   120
      TabIndex        =   6
      Top             =   465
      Width           =   7200
      Begin VB.TextBox txt_descripcion_2 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   1095
         Width           =   5820
      End
      Begin VB.TextBox txt_ubicacion 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1485
         Width           =   2295
      End
      Begin VB.TextBox txt_descripcion_1 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   705
         Width           =   5820
      End
      Begin VB.TextBox txt_codigo 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   315
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción 2:"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción 1:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1575
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   390
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmetiquetas_ubicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imprimir_Click()
   If Trim(Me.txt_codigo) <> "" Then
      If Trim(Me.txt_descripcion_1) <> "" Then
         If Trim(Me.txt_ubicacion) <> "" Then
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CreateTextFile(App.Path + "\etiquetas.txt", True)
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
            a.writeline ("A756,50,1,4,2,1,N,""UBICACION""")
            a.writeline ("A650,50,1,5,1,1,N,""" + Me.txt_ubicacion + """")
            a.writeline ("A550,50,1,4,2,1,N,""CODIGO""")
            a.writeline ("A450,50,1,5,1,1,N,""" + Me.txt_codigo + """")
            a.writeline ("A350,50,1,4,2,1,N,""DESCRIPCION""")
            a.writeline ("A250,50,1,5,1,1,N,""" + Me.txt_descripcion_1 + """")
            a.writeline ("A150,50,1,5,1,1,N,""" + Me.txt_descripcion_2 + """")
            a.writeline ("P1")
            a.Close
            Open (App.Path & "\etiquetas.bat") For Output As #2
            var_Archivo = App.Path & "\etiquetas.bat"
            Print #2, "copy " + App.Path + "\etiquetas.txt lpt1"
            Close #2
            X = Shell(var_Archivo, vbHide)
         Else
            MsgBox "Debe de indicar la ubicación del artículo", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Debe de indicar la descripción del artículo", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Debe de indicar el código del artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      Me.txt_ubicacion = ""
      Me.txt_descripcion_1 = ""
      Me.txt_descripcion_2 = ""
      rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_aRTICULO_ID = '" + IIf(IsNull(rs!VCHA_ART_aRTICULO_ID), "", rs!VCHA_ART_aRTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            Me.txt_descripcion_1 = IIf(IsNull(rsaux!VCHA_ART_NOMBRE_ESPAÑOL), "", rsaux!VCHA_ART_NOMBRE_ESPAÑOL)
            rsaux2.Open "SELECT * FROM TB_UBICACIONES_almacen WHERE VCHA_aRT_aRTICULO_ID = '" + IIf(IsNull(rs!VCHA_ART_aRTICULO_ID), "", rs!VCHA_ART_aRTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               Me.txt_ubicacion = IIf(IsNull(rsaux2!vcha_ubi_ubicacion_1), "", rsaux2!vcha_ubi_ubicacion_1)
            Else
               Me.txt_ubicacion = ""
            End If
            rsaux2.Close
         Else
            MsgBox "El código no existe", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      Else
         MsgBox "El código no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_descripcion_1 = ""
      Me.txt_descripcion_2 = ""
      Me.txt_ubicacion = ""
   End If
End Sub

Private Sub txt_descripcion_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descripcion_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
