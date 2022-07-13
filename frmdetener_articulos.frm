VERSION 5.00
Begin VB.Form frmdetener_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de archivo para detener artículos"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3900
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   5
         Top             =   510
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   3330
         Pattern         =   "*.xls"
         TabIndex        =   4
         Top             =   510
         Width           =   3075
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   90
         TabIndex        =   3
         Top             =   930
         Width           =   3150
      End
      Begin VB.CommandButton cmd_buscar_pedido 
         Caption         =   "Detener artículos"
         Height          =   465
         Left            =   3330
         TabIndex        =   2
         Top             =   2805
         Width           =   3060
      End
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3390
         Width           =   6315
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   6
         Top             =   120
         Width           =   6465
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
      Height          =   195
      Left            =   3720
      TabIndex        =   7
      Top             =   4050
      Width           =   540
   End
End
Attribute VB_Name = "frmdetener_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_buscar_pedido_Click()
   On Error GoTo salir:
   If Me.txt_ruta <> "" Then
      
      strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & Me.txt_ruta
      rsaux2.Open "SELECT codigo, detenido FROM [HOJA1$]", strConnectionString
      If Not rsaux2.EOF Then
         var_si = MsgBox("¿Desea cambiar el estatus detenido para los artículos indicados en el archivo?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el estatus detenido de los artículos en el archivo", vbYesNo, "ATENCION")
            If var_si = 6 Then
               While Not rsaux2.EOF
                     rsaux.Open "update tb_articulos set inte_Art_detenido = " + IIf(IsNull(rsaux2!detenido), 0, rsaux2!detenido) + " where vcha_Art_articulo_id = '" + IIf(IsNull(rsaux2!vcha_Art_articulo_id), "", rsaux2!vcha_Art_articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux2.MoveNext
               Wend
               MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
            End If
         End If
      End If
      rsaux2.Close
      
   Else
      MsgBox "No se a seleccionado un archivo", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
End Sub

Private Sub Dir1_Change()
   Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error GoTo salir:
   Me.Dir1.Path = Me.Drive1.Drive
   Me.Dir1.Refresh
   Exit Sub
salir:
   MsgBox "Unidad incorrecta"
   Me.Drive1.Drive = "c:"
End Sub

Private Sub File1_Click()
   If CStr(Me.Dir1.Path) = "C:\" Or CStr(Me.Dir1.Path) = "c:\" Then
      Me.txt_ruta = CStr(Me.Dir1.Path) + Me.File1.FileName
   Else
      Me.txt_ruta = CStr(Me.Dir1.Path) + "\" + Me.File1.FileName
   End If
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 2300
End Sub
