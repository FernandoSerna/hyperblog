VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignar_prioridad_rutas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar prioridad a rutas"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_clientes 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   750
      Picture         =   "frmoracle_asignar_prioridad_rutas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Clientes"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   120
      TabIndex        =   5
      Top             =   405
      Width           =   8295
      Begin VB.TextBox txt_prioridad 
         Height          =   390
         Left            =   945
         TabIndex        =   12
         Top             =   1035
         Width           =   720
      End
      Begin VB.TextBox txt_nombre 
         Height          =   420
         Left            =   945
         TabIndex        =   10
         Top             =   585
         Width           =   7230
      End
      Begin VB.TextBox txt_clave 
         Height          =   390
         Left            =   945
         TabIndex        =   8
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prioridad:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   705
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   263
         Width           =   450
      End
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   420
      Picture         =   "frmoracle_asignar_prioridad_rutas.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton com_nuevo_orden 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   90
      Picture         =   "frmoracle_asignar_prioridad_rutas.frx":05F4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame6 
      Height          =   75
      Left            =   75
      TabIndex        =   2
      Top             =   330
      Width           =   8340
   End
   Begin VB.Frame Frame2 
      Height          =   5280
      Left            =   120
      TabIndex        =   0
      Top             =   1950
      Width           =   8310
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   4725
         Left            =   60
         TabIndex        =   1
         Top             =   495
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8334
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Ruta"
            Object.Width           =   10142
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Prioridad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Rutas"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   8205
      End
   End
End
Attribute VB_Name = "frmoracle_asignar_prioridad_rutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_clientes_Click()
   var_ruta_cliente = Me.lv_rutas.selectedItem
   var_nombre_ruta_cliente = Me.lv_rutas.selectedItem.SubItems(1)
   frmoracle_rutas_clientes.Show 1
End Sub

Private Sub com_guardar_Click()
   If IsNumeric(Me.txt_prioridad) Then
      var_si = MsgBox("¿Desea guardar los cambios?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "UPDATE TB_ORACLE_RUTAS SET PRIORIDAD = " + Me.txt_prioridad + " WHERE CLAVE = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_rutas.selectedItem.SubItems(2) = Me.txt_prioridad
         Me.lv_rutas.SetFocus
      End If
   Else
      MsgBox "Número de prioridad incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 1500
   rs.Open "select * from tb_oracle_rutas", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_rutas.ListItems.Add(, , rs!CLAVE)
         list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_RUTA), "", rs!NOMBRE_RUTA)
         list_item.SubItems(2) = IIf(IsNull(rs!prioridad), "", rs!prioridad)
         rs.MoveNext
      Wend
      rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_rutas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_rutas, ColumnHeader)
End Sub

Private Sub lv_rutas_GotFocus()
   Me.txt_clave = Me.lv_rutas.selectedItem
   Me.txt_nombre = Me.lv_rutas.selectedItem.SubItems(1)
   Me.txt_prioridad = Me.lv_rutas.selectedItem.SubItems(2)
End Sub

Private Sub lv_rutas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.txt_clave = Me.lv_rutas.selectedItem
   Me.txt_nombre = Me.lv_rutas.selectedItem.SubItems(1)
   Me.txt_prioridad = Me.lv_rutas.selectedItem.SubItems(2)
End Sub

Private Sub lv_rutas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_prioridad.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre.SetFocus
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
      Me.txt_prioridad.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_prioridad_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.com_guardar.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
 End Sub
