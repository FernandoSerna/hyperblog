VERSION 5.00
Begin VB.Form frmoracle_cargar_cajas_houston 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar cajas a Houston"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1800
      Left            =   135
      TabIndex        =   0
      Top             =   -15
      Width           =   4410
      Begin VB.TextBox txtfactura 
         Height          =   570
         Left            =   150
         TabIndex        =   2
         Top             =   285
         Width           =   4080
      End
      Begin VB.CommandButton cmdmandar_informacion 
         Caption         =   "Mandar Información"
         Height          =   675
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmoracle_cargar_cajas_houston"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmandar_informacion_Click()
   If Trim(Me.txtfactura) <> "" Then
      
   End If
End Sub

Private Sub txtfactura_Change()
x = "SELECT * FROM RA_CUSTOMER_TRX_LINES_ALL WHERE SALES_ORDER = " + CStr(var_numero_pedido)
End Sub
