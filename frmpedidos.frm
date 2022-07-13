VERSION 5.00
Begin VB.Form frmpedidos 
   Caption         =   "Pedidos"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   3
      Top             =   1050
      Width           =   10920
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -15
      TabIndex        =   2
      Top             =   405
      Width           =   10920
   End
   Begin VB.ComboBox cmb_tipopedidos 
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   690
      Width           =   4230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Pedido:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   735
      Width           =   1140
   End
End
Attribute VB_Name = "frmpedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    rs.Open "select * from tb_tipopedidos", cnn, adOpenDynamic, adLockBatchOptimistic
    Call RecsetToCombo(cmb_tipopedidos.hwnd, rs, 1)
    rs.Close
End Sub
