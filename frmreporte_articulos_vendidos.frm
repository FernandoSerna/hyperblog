VERSION 5.00
Begin VB.Form frmreporte_articulos_vendidos 
   Caption         =   "Reporte de Artículos "
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11580
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_articulos_vendidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Generar Reporte "
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmreporte_articulos_vendidos.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_articulos_vendidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()

End Sub
