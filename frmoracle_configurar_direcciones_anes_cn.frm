VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_configurar_direcciones_anes_cn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de clientes para Carta Porte"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Frame Frame1 
      Height          =   3480
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   9375
      Begin VB.TextBox txt_cp 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3000
         Width           =   1500
      End
      Begin VB.TextBox txt_distancia_AGS 
         Height          =   315
         Left            =   4350
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1800
         Width           =   1500
      End
      Begin VB.TextBox txt_distancia_CDMX 
         Height          =   315
         Left            =   4350
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2160
         Width           =   1515
      End
      Begin VB.TextBox txt_distancia_MTY 
         Height          =   315
         Left            =   4350
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2535
         Width           =   1500
      End
      Begin VB.TextBox txt_calle 
         Height          =   315
         Left            =   1530
         TabIndex        =   20
         Top             =   840
         Width           =   4155
      End
      Begin VB.TextBox txt_clave 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   11
         Top             =   160
         Width           =   1500
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1530
         TabIndex        =   10
         Top             =   500
         Width           =   4155
      End
      Begin VB.TextBox txt_numero_externo 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1200
         Width           =   1545
      End
      Begin VB.TextBox txt_numero_interno 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox txt_colonia 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1920
         Width           =   1500
      End
      Begin VB.TextBox txt_municipio 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2280
         Width           =   1515
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2655
         Width           =   1500
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   195
         Index           =   11
         Left            =   300
         TabIndex        =   29
         Top             =   3045
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Distanica AGS:"
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   27
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Distancia CDMX:"
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   26
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Distancia MTY:"
         Height          =   195
         Index           =   8
         Left            =   3120
         TabIndex        =   25
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Calle:"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   21
         Top             =   885
         Width           =   390
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   18
         Top             =   180
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   540
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Número Externo:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   16
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Número Interno:"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   15
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   14
         Top             =   1980
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   13
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   12
         Top             =   2700
         Width           =   540
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_configurar_direcciones_anes_cn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      Picture         =   "frmoracle_configurar_direcciones_anes_cn.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   885
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ListView lv_clientes 
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4868
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Calle"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "numero_externo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Numero_interno"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Colonia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Municipio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Estado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Distancia_AGS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Distancia_CDMX"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Distancia_MTY"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Codigo Postal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "configvehicular"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "placavm"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "aniomodelovm"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "subtiporem"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "placa"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_configurar_direcciones_anes_cn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If Me.txt_clave.Text <> "" Then
      If IsNumeric(Me.txt_distancia_AGS) Then
         If IsNumeric(Me.txt_distancia_CDMX) Then
            If IsNumeric(Me.txt_distancia_MTY) Then
               rsaux.Open "select * from xxvia_tb_anes_carta_porte where clave = '" + Me.txt_clave + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rs.Open "update xxvia_tb_anes_carta_porte set nombre = '" + Me.txt_nombre + "', calle = '" + Me.txt_calle + "', NUMERO_EXTERIOR = '" + Me.txt_numero_externo + "', NUMERO_INTERIOR = '" + Me.txt_numero_interno + "', colonia = '" + Me.txt_colonia + "', municipio = '" + Me.txt_municipio + "', estado = '" + Me.txt_estado + "', distancia_ags = '" + Me.txt_distancia_AGS + "', distancia_cdm = " + Me.txt_distancia_CDMX + ", distancia_mty = " + Me.txt_distancia_MTY + ", CODIGO_POSTAL = '" + Me.txt_cp + "' where clave = '" + Me.txt_clave + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Me.lv_clientes.ListItems.Clear
                  rs.Open "select * from xxvia_tb_anes_carta_porte where clave not like 'C0%' and clave not like '00%' order by clave", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        Set list_item = lv_clientes.ListItems.Add(, , rs!clave)
                        list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
                        list_item.SubItems(2) = IIf(IsNull(rs!calle), "", rs!calle)
                        list_item.SubItems(3) = IIf(IsNull(rs!numero_exterior), "", rs!numero_exterior)
                        list_item.SubItems(4) = IIf(IsNull(rs!numero_interior), "", rs!numero_interior)
                        list_item.SubItems(5) = IIf(IsNull(rs!colonia), "", rs!colonia)
                        list_item.SubItems(6) = IIf(IsNull(rs!municipio), "", rs!municipio)
                        list_item.SubItems(7) = IIf(IsNull(rs!estado), "", rs!estado)
                        list_item.SubItems(8) = IIf(IsNull(rs!DISTANCIA_AGS), "0", rs!DISTANCIA_AGS)
                        list_item.SubItems(9) = IIf(IsNull(rs!DISTANCIA_CDM), "0", rs!DISTANCIA_CDM)
                        list_item.SubItems(10) = IIf(IsNull(rs!DISTANCIA_MTY), "0", rs!DISTANCIA_MTY)
                        list_item.SubItems(11) = IIf(IsNull(rs!codigo_postal), "", rs!codigo_postal)
                        rs.MoveNext
                  Wend
                  rs.Close
                  MsgBox "Se a actualizado el cliente.", vbOKOnly, "ATENCION"
               Else
                  rs.Open "insert into  xxvia_tb_anes_carta_porte (clave, nombre, calle, numero_exterior, numero_interior, colonia, municipio, estado, distancia_ags, distancia_cdm, distancia_mty, codigo_postal) values ('" + Me.txt_clave + "', '" + Me.txt_nombre + "', '" + Me.txt_calle + "', '" + Me.txt_numero_externo + "', '" + Me.txt_numero_interno + "', '" + Me.txt_colonia + "', '" + Me.txt_municipio + "', '" + Me.txt_estado + "', '" + Me.txt_distancia_AGS + "', " + Me.txt_distancia_CDMX + ", " + Me.txt_distancia_MTY + ", '" + Me.txt_cp + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Me.lv_clientes.ListItems.Clear
                  rs.Open "select * from xxvia_tb_anes_carta_porte where clave not like 'C0%' and clave not like '00%' order by clave", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        Set list_item = lv_clientes.ListItems.Add(, , rs!clave)
                        list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
                        list_item.SubItems(2) = IIf(IsNull(rs!calle), "", rs!calle)
                        list_item.SubItems(3) = IIf(IsNull(rs!numero_exterior), "", rs!numero_exterior)
                        list_item.SubItems(4) = IIf(IsNull(rs!numero_interior), "", rs!numero_interior)
                        list_item.SubItems(5) = IIf(IsNull(rs!colonia), "", rs!colonia)
                        list_item.SubItems(6) = IIf(IsNull(rs!municipio), "", rs!municipio)
                        list_item.SubItems(7) = IIf(IsNull(rs!estado), "", rs!estado)
                        list_item.SubItems(8) = IIf(IsNull(rs!DISTANCIA_AGS), "0", rs!DISTANCIA_AGS)
                        list_item.SubItems(9) = IIf(IsNull(rs!DISTANCIA_CDM), "0", rs!DISTANCIA_CDM)
                        list_item.SubItems(10) = IIf(IsNull(rs!DISTANCIA_MTY), "0", rs!DISTANCIA_MTY)
                        list_item.SubItems(11) = IIf(IsNull(rs!codigo_postal), "", rs!codigo_postal)
                        rs.MoveNext
                  Wend
                  rs.Close
                  
                  MsgBox "Se a ingresado el cliente", vbOKOnly, "ATENCION"
               End If
               MsgBox "Se ha actualizado el registro", vbOKOnly, "ATENCION"
            Else
               MsgBox "La distancia a Monterrey es incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La distancia en CDMX es incorrecta"
         End If
      Else
         MsgBox "La distancia en Aguascalientes es incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_calle = ""
   Me.txt_clave = ""
   Me.txt_colonia = ""
   Me.txt_cp = ""
   Me.txt_distancia_AGS = ""
   Me.txt_distancia_CDMX = ""
   Me.txt_distancia_MTY = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_nombre = ""
   Me.txt_numero_externo = ""
   Me.txt_numero_interno = ""
   Me.txt_clave.SetFocus
   
End Sub

Private Sub Form_Load()
   rs.Open "select * from xxvia_tb_anes_carta_porte where clave not like 'C0%' and clave not like '00%' order by clave", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_clientes.ListItems.Add(, , rs!clave)
         list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!calle), "", rs!calle)
         list_item.SubItems(3) = IIf(IsNull(rs!numero_exterior), "", rs!numero_exterior)
         list_item.SubItems(4) = IIf(IsNull(rs!numero_interior), "", rs!numero_interior)
         list_item.SubItems(5) = IIf(IsNull(rs!colonia), "", rs!colonia)
         list_item.SubItems(6) = IIf(IsNull(rs!municipio), "", rs!municipio)
         list_item.SubItems(7) = IIf(IsNull(rs!estado), "", rs!estado)
         list_item.SubItems(8) = IIf(IsNull(rs!DISTANCIA_AGS), "0", rs!DISTANCIA_AGS)
         list_item.SubItems(9) = IIf(IsNull(rs!DISTANCIA_CDM), "0", rs!DISTANCIA_CDM)
         list_item.SubItems(10) = IIf(IsNull(rs!DISTANCIA_MTY), "0", rs!DISTANCIA_MTY)
         list_item.SubItems(11) = IIf(IsNull(rs!codigo_postal), "0", rs!codigo_postal)
         rs.MoveNext
      Wend
      rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_clientes, ColumnHeader)

End Sub

Private Sub lv_clientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_clientes.ListItems.Count > 0 Then
      Me.txt_clave = Me.lv_clientes.selectedItem
      Me.txt_nombre = Me.lv_clientes.selectedItem.SubItems(1)
      Me.txt_calle = Me.lv_clientes.selectedItem.SubItems(2)
      Me.txt_numero_externo = Me.lv_clientes.selectedItem.SubItems(3)
      Me.txt_numero_interno = Me.lv_clientes.selectedItem.SubItems(4)
      Me.txt_colonia = Me.lv_clientes.selectedItem.SubItems(5)
      Me.txt_municipio = Me.lv_clientes.selectedItem.SubItems(6)
      Me.txt_estado = Me.lv_clientes.selectedItem.SubItems(7)
      Me.txt_distancia_AGS = Me.lv_clientes.selectedItem.SubItems(8)
      Me.txt_distancia_CDMX = Me.lv_clientes.selectedItem.SubItems(9)
      Me.txt_distancia_MTY = Me.lv_clientes.selectedItem.SubItems(10)
      Me.txt_cp = Me.lv_clientes.selectedItem.SubItems(11)
   End If

End Sub

Private Sub txt_calle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_numero_externo.SetFocus
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre.SetFocus
   End If
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_municipio.SetFocus
   End If
End Sub

Private Sub txt_cp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_distancia_AGS.SetFocus
   End If
End Sub

Private Sub txt_distancia_AGS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_distancia_CDMX.SetFocus
   End If
End Sub

Private Sub txt_distancia_CDMX_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_distancia_MTY.SetFocus
   End If
End Sub

Private Sub txt_distancia_MTY_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cp.SetFocus
   End If
End Sub

Private Sub txt_municipio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_estado.SetFocus
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_calle.SetFocus
   End If
End Sub

Private Sub txt_numero_externo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_numero_interno.SetFocus
   End If
End Sub

Private Sub txt_numero_interno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_colonia.SetFocus
   End If
End Sub
