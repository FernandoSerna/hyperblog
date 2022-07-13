VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_bitacora_aduana_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bitacora de bultos leidos por aduana"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5130
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   9780
      Begin MSComctlLib.ListView lv_bultos 
         Height          =   4500
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   7938
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bulto"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Maquina"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DVR"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Puerto"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Lectura"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Información del pedido"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   135
         Width           =   9690
      End
   End
   Begin VB.TextBox txt_pedido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label lbl_cliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   9495
   End
   Begin VB.Label lbl_embarque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   1005
   End
End
Attribute VB_Name = "frmoracle_bitacora_aduana_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Top = 0
  Left = 1000
  Me.lbl_cliente = ""
  Me.lbl_embarque = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_pedido_Change()
   Me.lbl_cliente = ""
   Me.lbl_embarque = ""
   Me.lv_bultos.ListItems.Clear
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         rs.Open "SELECT * FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.lbl_embarque = "EMBARQUE: " + CStr(rs!Embarque)
            Me.lbl_cliente = "CLIENTE: " + rs!NOMBRE_AGENTE
            rsaux.Open "select * FROM TB_ORACLE_BITACORA_CAJAS_ADUANA where PEDIDO = '" + Me.txt_pedido + "' ORDER BY CAJA", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               While Not rsaux.EOF
                     Set list_item = lv_bultos.ListItems.Add(, , rsaux!Caja)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!MAQUINA), "", rsaux!MAQUINA)
                     list_item.SubItems(2) = IIf(IsNull(rsaux!DVR), "", rsaux!DVR)
                     list_item.SubItems(3) = IIf(IsNull(rsaux!PUERTO), "", rsaux!PUERTO)
                     list_item.SubItems(4) = IIf(IsNull(rsaux!Fecha), "", rsaux!Fecha)

                     rsaux.MoveNext
               Wend
            
            Else
               MsgBox "No existe bitacora para el pedido seleccionado", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         Else
            MsgBox "El pedido no existe.", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Pedido incorrecto.", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
