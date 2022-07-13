VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmbloques 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmbloques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmd_eliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmd_aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin TabDlg.SSTab tabBloques 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Bloques del Sistema"
      TabPicture(0)   =   "frmbloques.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lab_fecha_bloques"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtBloques(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtBloques(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtBloques 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtBloques 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lab_fecha_bloques 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   360
         Picture         =   "frmbloques.frx":08E6
         Stretch         =   -1  'True
         Top             =   600
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   90
         Index           =   0
         Left            =   120
         Picture         =   "frmbloques.frx":0E70
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   4995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Bloque"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Clave del Bloque"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmbloques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean

Private Sub cmd_aceptar_Click()

Dim ok As Boolean, var_llave_siguiente As String

Set Tb_Bloques = New Tb_Bloques
    
    ok = True
    If txtBloques(0) <> "" And txtBloques(1) <> "" Then
        If var_hubo_cambios Then
            ok = Tb_Bloques.Anadir(txtBloques(0), txtBloques(1), lab_fecha_bloques, fun_NombreUsuario, fun_NombrePc, var_numero_planta, "")
            If ok Then
                pro_actualiza_ListView
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                Unload Me
            Else
                MsgBox "No se puede grabar registro: " + Tb_Bloques.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set Tb_Bloques = Nothing: var_hubo_cambios = False

End Sub

Private Sub cmd_cancelar_Click()

    Unload Me

End Sub

Private Sub cmd_eliminar_Click()

Dim var_llave_usuarios As String

Set Tb_Bloques = New Tb_Bloques

    
    ok = True
    If txtBloques(0) <> "" And txtBloques(1) <> "" Then
        var_llave_usuarios = Tb_Bloques.Obtener_llave(cnn, rs, "vcha_blo_bloque_id", txtBloques(0))
        If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = Tb_Bloques.Eliminar(var_llave_usuarios)
        Else
            ok = False
        End If
        If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
                frmmenus.lvbloques.ListItems.Remove (frmmenus.lvbloques.SelectedItem.Index)
            Unload Me
        Else
            MsgBox "No se puede Eliminar registro: " + Tb_Bloques.MensajeError, vbOKOnly + vbCritical, "ATENCION"
        End If
    End If
    
    
Set lvbloques = Nothing

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Call pro_enfoque(KeyAscii)
    
End Sub

Private Sub Form_Load()

Set Tb_Bloques = New Tb_Bloques
    txtBloques(0) = Format(Tb_Bloques.Siguiente(cnn, rs), "00")
Set Tb_Bloques = Nothing
    lab_fecha_bloques = Format(Date, "dd/mm/yyyy")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmmenus.lvbloques.ListItems.Item(frmmenus.lvbloques.SelectedItem.Index).Checked = False
    
End Sub

Private Sub txtBloques_GotFocus(Index As Integer)

    txtBloques(Index).BackColor = vbVioletBright
    txtBloques(Index).SelStart = 0
    txtBloques(Index).SelLength = Len(txtBloques(Index).Text)

End Sub

Private Sub txtBloques_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    var_hubo_cambios = True
End Sub

Private Sub txtBloques_LostFocus(Index As Integer)

    txtBloques(Index).BackColor = &H80000005
    
End Sub




'++++++++++++++++++++ ACTUALIZA CAMBIOS HECHOS EN LISTAS +++++++++++++++++++++++

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
   With frmmenus
        If var_modifica_registro = False Then
            Set list_item = .lvbloques.ListItems.Add(, , txtBloques(1)): list_item.SmallIcon = 3
            list_item.SubItems(1) = txtBloques(0)

        Else
            .lvbloques.ListItems.Item(.lvbloques.SelectedItem.Index).Checked = False
            .lvbloques.ListItems.Item(.lvbloques.SelectedItem.Index) = txtBloques(1)
            .lvbloques.ListItems.Item(.lvbloques.SelectedItem.Index).ListSubItems(1) = txtBloques(0)
        End If
   End With
End Sub
