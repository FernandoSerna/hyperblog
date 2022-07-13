VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignar_ruta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar ruta"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6420
      Picture         =   "frmoracle_asignar_ruta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_asignar_ruta.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   4305
      Left            =   15
      TabIndex        =   5
      Top             =   1185
      Width           =   6735
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   3705
         Left            =   75
         TabIndex        =   6
         Top             =   495
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   6535
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Ruta"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Rutas"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   45
         TabIndex        =   7
         Top             =   135
         Width           =   6645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6780
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   15
      TabIndex        =   0
      Top             =   450
      Width           =   6735
      Begin VB.TextBox txt_nombre 
         Height          =   360
         Left            =   1425
         TabIndex        =   3
         Top             =   225
         Width           =   5190
      End
      Begin VB.TextBox txt_clave 
         Height          =   360
         Left            =   660
         TabIndex        =   1
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   285
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmoracle_asignar_ruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If Me.txt_clave <> "" Then
      If var_tipo_asigna_ruta = 1 Then
         rs.Open "select * from TB_ORACLE_RUTAS_EMBARQUES WHERE EMBARQUE = " + CStr(var_embarque_ruta), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "UPDATE TB_ORACLE_RUTAS_EMBARQUES SET RUTA = '" + Me.txt_clave + "' WHERE EMBARQUE = " + CStr(var_embarque_ruta), cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
         Else
            rsaux.Open "INSERT TB_ORACLE_RUTAS_EMBARQUES (EMBARQUE, RUTA) VALUES ('" + CStr(var_embarque_ruta) + "','" + Me.txt_clave + "')", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         If var_tipo_asigna_ruta = 2 Then
            var_ruta_distribucion = Me.txt_clave
            Unload Me
         End If
      End If
   Else
      MsgBox "No se a seleccionado una ruta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 200
   Left = 2000
   rs.Open "SELECT * FROM XXVIA_TB_RUTAS_DISTRIBUCION", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_rutas.ListItems.Add(, , rs!ruta)
         list_item.SubItems(1) = IIf(IsNull(rs!nombre_ruta), 0, rs!nombre_ruta)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub lv_rutas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.txt_clave = Me.lv_rutas.selectedItem
      Me.txt_nombre = Me.lv_rutas.selectedItem.SubItems(1)
   End If
End Sub

Private Sub lv_rutas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
