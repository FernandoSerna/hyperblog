VERSION 5.00
Begin VB.Form frmoracle_lectura_guia 
   BorderStyle     =   0  'None
   Caption         =   "Lectura de Guia"
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7245
      Begin VB.TextBox txt_codigo 
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
         Left            =   900
         TabIndex        =   1
         Top             =   180
         Width           =   6225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
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
         Left            =   105
         TabIndex        =   2
         Top             =   330
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmoracle_lectura_guia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
