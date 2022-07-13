VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmclonacionagrupadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clonación de Agrupadores"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmclonacionagrupadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6645
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmclonacionagrupadores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmclonacionagrupadores.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6180
      Picture         =   "frmclonacionagrupadores.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos del Nuevo Agrupador "
      Height          =   1080
      Left            =   105
      TabIndex        =   1
      Top             =   1290
      Width           =   6480
      Begin VB.TextBox txt_clonasionagrupadores 
         Height          =   315
         Index           =   2
         Left            =   810
         MaxLength       =   50
         TabIndex        =   7
         Top             =   615
         Width           =   5505
      End
      Begin VB.TextBox txt_clonasionagrupadores 
         Height          =   315
         Index           =   1
         Left            =   810
         MaxLength       =   50
         TabIndex        =   5
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   615
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   450
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2835
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":10D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":19B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":228C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":2828
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":3102
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":39DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":42B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":45D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":48EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":4E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":51A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclonacionagrupadores.frx":52B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   " Agrupador "
      Height          =   795
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   6465
      Begin VB.ComboBox cmbclonasionagrupadores 
         Height          =   315
         Index           =   0
         Left            =   810
         TabIndex        =   4
         Top             =   315
         Width           =   5535
      End
      Begin VB.TextBox txt_clonasionagrupadores 
         Height          =   285
         Index           =   0
         Left            =   915
         TabIndex        =   2
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   345
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmclonacionagrupadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbclonasionagrupadores_Click(Index As Integer)
   txt_clonasionagrupadores(0) = Obtener_llave(cnn, rs, "TB_AGRUPADORES", "VCHA_AGR_NOMBRE", cmbclonasionagrupadores(0), 1, "T")
End Sub

Private Sub cmd_nuevo_Click()
        rs.Open " select * from tb_agrupadores where vcha_agr_agrupador_id = '" + txt_clonasionagrupadores(1) + "'", cnn, adOpenDynamic, adLockOptimistic
        If rs.EOF Then
           rs.Close
           rs.Open " select * from tb_agrupadores where vcha_agr_nombre = '" + txt_clonasionagrupadores(2) + "'", cnn, adOpenDynamic, adLockOptimistic
           If rs.EOF Then
              tipo = "C"
              familia = frmagrupadores2.lv_familia_agrupadores.selectedItem
              agrupador = txt_clonasionagrupadores(1)
              nombreagrupador = txt_clonasionagrupadores(2)
              rsaux2.Open "Insert into tb_agrupadores (VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_AGR_NOMBRE, VCHA_AGR_TIPO) values ('" + familia + "', '" + agrupador + "', '" + nombreagrupador + "', '" + tipo + "')", cnn, adOpenDynamic, adLockOptimistic
              rs.Close
              rs.Open "select * from tb_detalle_agrupadores where vcha_agr_agrupador_id = '" + txt_clonasionagrupadores(0) + "'", cnn, adOpenDynamic, adLockOptimistic
              While Not rs.EOF
                 If rs(1).Value = 1 Then
                    rsaux2.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO, VCHA_ART_ARTICULO_ID, VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values  ('" + agrupador + "', '" + Str(rs(1).Value) + "', '" + rs(2).Value + "', Null ,  Null ,  Null , Null )", cnn, adOpenDynamic, adLockOptimistic
                 End If
                 If rs(1).Value = 2 Then
                    rsaux2.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO, VCHA_ART_ARTICULO_ID, VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values  ('" + agrupador + "', '" + Str(rs(1).Value) + "', Null , '" + rs(3).Value + "',  Null ,  Null ,  Null )", cnn, adOpenDynamic, adLockOptimistic
                 End If
                 If rs(1).Value = 3 Then
                    rsaux2.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO, VCHA_ART_ARTICULO_ID, VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values  ('" + agrupador + "', '" + Str(rs(1).Value) + "', null , '" + rs(3).Value + "','" + rs(4).Value + "', null, null)", cnn, adOpenDynamic, adLockOptimistic
                 End If
                 If rs(1).Value = 4 Then
                    rsaux2.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO, VCHA_ART_ARTICULO_ID, VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values  ('" + agrupador + "', '" + Str(rs(1).Value) + "', null, null , null, '" + rs(5).Value + "', null)", cnn, adOpenDynamic, adLockOptimistic
                 End If
                 If rs(1).Value = 5 Then
                    rsaux2.Open "insert into tb_detalle_agrupadores (VCHA_AGR_AGRUPADOR_ID, INTE_DEA_TIPO, VCHA_ART_ARTICULO_ID, VCHA_LIN_LINEA_ID, VCHA_SLI_SUBLINEA_ID, VCHA_PRO_PRODUCTO_ID, VCHA_TAR_TIPO_ARTICULO_ID) values  ('" + agrupador + "', '" + Str(rs(1).Value) + "', null, null , null, null, '" + rs(6).Value + "')", cnn, adOpenDynamic, adLockOptimistic
                 End If
                 rs.MoveNext
              Wend
              rs.Close
              MsgBox "Se a terminado el proceso de clonasión del agrupador", vbOKOnly, "ATENCION"
           Else
              rs.Close
              MsgBox "Ya existe un agrupador con el nombre " + txt_clonasionagrupadores(2) + ", favor de seleccionar otro", vbOKOnly, "ATENCION"
           End If
        Else
           rs.Close
           MsgBox "Ya existe un agrupador con la clave " + txt_clonasionagrupadores(1) + ", favor de seleccionar otra", vbOKOnly, "ATENCION"
        End If
        
        

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   rs.Open "select * from tb_agrupadores order by vcha_agr_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmbclonasionagrupadores(0).hwnd, rs, 2)
   rs.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_clonacionagrupadores)
End Sub

Private Sub txt_clonasionagrupadores_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
