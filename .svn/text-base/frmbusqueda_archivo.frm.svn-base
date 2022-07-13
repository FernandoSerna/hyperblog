VERSION 5.00
Begin VB.Form frmbusqueda_archivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         Picture         =   "frmbusqueda_archivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   3255
         Width           =   330
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   4
         Top             =   555
         Width           =   3165
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   3330
         Pattern         =   "*.xls"
         TabIndex        =   3
         Top             =   525
         Width           =   3075
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   90
         TabIndex        =   2
         Top             =   1020
         Width           =   3150
      End
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3255
         Width           =   5805
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   6465
      End
   End
End
Attribute VB_Name = "frmbusqueda_archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   var_archivo_buscar = Me.txt_ruta
   Unload Me
End Sub

Private Sub Dir1_Change()
   Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error GoTo salir:
   Me.Dir1.Path = Me.Drive1.Drive
   Me.Dir1.Refresh
   Exit Sub
salir:
   MsgBox "Unidad incorrecta"
   Me.Drive1.Drive = "c:"
End Sub

Private Sub File1_Click()
   If CStr(Me.Dir1.Path) = "C:\" Or CStr(Me.Dir1.Path) = "c:\" Then
      Me.txt_ruta = CStr(Me.Dir1.Path) + Me.File1.FileName
   Else
      Me.txt_ruta = CStr(Me.Dir1.Path) + "\" + Me.File1.FileName
   End If
End Sub
