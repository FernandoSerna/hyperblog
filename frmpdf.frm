VERSION 5.00
Object = "{CA8A9783-280D-11CF-A24D-444553540000}#1.3#0"; "pdf.ocx"
Begin VB.Form frmpdf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PdfLib.Pdf Pdf1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _Version        =   327680
      _ExtentX        =   20532
      _ExtentY        =   12938
      _StockProps     =   0
      SRC             =   ""
   End
End
Attribute VB_Name = "frmpdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Top = 0
   Left = 0
   Me.Caption = "Visor de facturas electrónicas"
   Me.Pdf1.src = var_ruta_factura_pdf
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub

