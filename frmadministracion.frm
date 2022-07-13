VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmadministracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento a Empresas"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmadministracion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   5280
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   13
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Eliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Empresas"
      TabPicture(0)   =   "frmadministracion.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "labfecha_alta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Image1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Image2(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Image2(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtempresa(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtempresa(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtempresa(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtempresa(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtempresa(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtempresa(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtempresa(6)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtempresa(7)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   20
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   6
         Left            =   3240
         TabIndex        =   16
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   15
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   14
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtempresa 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   90
         Index           =   1
         Left            =   240
         Picture         =   "frmadministracion.frx":08E6
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   4875
      End
      Begin VB.Image Image2 
         Height          =   90
         Index           =   0
         Left            =   240
         Picture         =   "frmadministracion.frx":0D23
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   4875
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmadministracion.frx":1160
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Alta"
         Height          =   195
         Left            =   2760
         TabIndex        =   22
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "RFC"
         Height          =   195
         Index           =   7
         Left            =   1035
         TabIndex        =   21
         Top             =   4440
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   19
         Top             =   4080
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CP"
         Height          =   195
         Index           =   5
         Left            =   1200
         TabIndex        =   18
         Top             =   4080
         Width           =   210
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Domicilio Fiscal"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label labfecha_alta 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   8
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Planta"
         Height          =   195
         Index           =   2
         Left            =   915
         TabIndex        =   7
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   600
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":146A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":1A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":1B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":2438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":2752
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   0
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":302C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmadministracion.frx":3906
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmadministracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_llave_auxiliar As String


Private Sub Aceptar_Click()
Dim ok As Boolean, var_llave_siguiente As String

Set Tb_plantas = New Tb_plantas
    
    ok = True
    If txtempresa(0) <> "" And txtempresa(1) <> "" And txtempresa(2) <> "" And txtempresa(3) <> "" Then
        ok = Tb_plantas.Anadir(txtempresa(0), txtempresa(1), txtempresa(2), txtempresa(3), txtempresa(4), txtempresa(5), txtempresa(6), txtempresa(7), labfecha_alta, fun_NombreUsuario, fun_NombrePc, 0)
        If ok Then
            pro_actualiza_ListView
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            Unload Me
        Else
            MsgBox "No se puede grabar registro: " + Tb_plantas.MensajeError, vbOKOnly + vbCritical, "ATENCION"
        End If
    End If
Set TB_EMPRESAS = Nothing

End Sub

Private Sub cancelar_Click()
    Unload Me
End Sub

Private Sub Eliminar_Click()
Dim var_llave_usuarios As String

Set Tb_plantas = New Tb_plantas
ok = True

    rs.Open "select * from TB_DETALLE where TB_DETALLE.BINT_PLA_PLANTA_ID = '" & txtempresa(0) & "'", cnn, adOpenForwardOnly, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        If txtempresa(0) <> "" And txtempresa(1) <> "" And txtempresa(2) <> "" And txtempresa(3) <> "" Then
            var_llave_usuarios = Str(Val(txtempresa(0)))
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = Tb_plantas.Eliminar(Trim(var_llave_usuarios))
            Else
                GoTo SALIR:
            End If
            If ok Then
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                frmmenus.lvplantas.ListItems.Remove (frmmenus.lvplantas.SelectedItem.Index)
                Unload Me
            Else
                MsgBox "No se puede grabar registro: " + Tb_plantas.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    Else
        SetTimer hwnd, NV_CLOSEMSGBOX, 1800, AddressOf TimerProc
        MsgBox "No se Puede Borrar Este Registro, Existen Dependencias", , "TRANSACCIONES [ AVISO ]"
        rs.Close
    End If

SALIR:
Set Tb_plantas = Nothing

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Call pro_enfoque(KeyAscii)
    
End Sub

Private Sub Form_Load()
Set Tb_plantas = New Tb_plantas
    txtempresa(0) = Format(Tb_plantas.Siguiente(cnn, rs), "00")
Set Tb_plantas = Nothing
    labfecha_alta = Format(Date, "dd/mm/yyyy")
End Sub






'++++++++++++++++++++ ACTUALIZA CAMBIOS HECHOS EN LISTAS +++++++++++++++++++++++


Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
   With frmmenus
        If var_modifica_registro = False Then
            Set list_item = .lvplantas.ListItems.Add(, , txtempresa(1)): list_item.SmallIcon = 2
            list_item.SubItems(1) = txtempresa(2)
            list_item.SubItems(2) = txtempresa(3)
            list_item.SubItems(3) = txtempresa(4)
            list_item.SubItems(4) = txtempresa(5)
            list_item.SubItems(5) = txtempresa(6)
            list_item.SubItems(6) = txtempresa(7)
            list_item.SubItems(7) = txtempresa(0)
        Else
            .lvplantas.ListItems.Item(.lvplantas.SelectedItem.Index).Checked = False
            .lvplantas.ListItems.Item(.lvplantas.SelectedItem.Index) = txtempresa(1)
            .lvplantas.SelectedItem.SubItems(1) = txtempresa(2)
            .lvplantas.SelectedItem.SubItems(2) = txtempresa(3)
            .lvplantas.SelectedItem.SubItems(3) = txtempresa(4)
            .lvplantas.SelectedItem.SubItems(4) = txtempresa(5)
            .lvplantas.SelectedItem.SubItems(5) = txtempresa(6)
            .lvplantas.SelectedItem.SubItems(6) = txtempresa(7)
            .lvplantas.SelectedItem.SubItems(7) = txtempresa(0)
        End If
   End With
End Sub





Private Sub Form_Unload(Cancel As Integer)
    frmmenus.lvplantas.ListItems.Item(frmmenus.lvplantas.SelectedItem.Index).Checked = False
End Sub

Private Sub txtempresa_GotFocus(Index As Integer)

    txtempresa(Index).BackColor = vbVioletBright
    txtempresa(Index).SelStart = 0
    txtempresa(Index).SelLength = Len(txtempresa(Index).Text)

End Sub

Private Sub txtempresa_LostFocus(Index As Integer)
    
    txtempresa(Index).BackColor = &H80000005

End Sub
