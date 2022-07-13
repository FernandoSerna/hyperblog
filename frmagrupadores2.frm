VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmagrupadores2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agrupadores"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "frmagrupadores2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11400
   Begin VB.Frame Frame5 
      Caption         =   "  Detalle del Agrupador "
      Height          =   3915
      Left            =   105
      TabIndex        =   34
      Top             =   3645
      Width           =   11400
      Begin VB.Frame Frame3 
         Height          =   3270
         Index           =   2
         Left            =   5805
         TabIndex        =   47
         Top             =   600
         Width           =   5565
         Begin MSComctlLib.ImageList ImageList1 
            Index           =   2
            Left            =   720
            Top             =   225
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   10
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":08CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":11A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":1A7E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":201A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":28F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":31CE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":3AA8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":3DC2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":40DC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":4678
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList icono_encabezado 
            Index           =   2
            Left            =   135
            Top             =   240
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
                  Picture         =   "frmagrupadores2.frx":4992
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":526C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lv_detalle_agrupadores 
            Height          =   3075
            Index           =   0
            Left            =   45
            TabIndex        =   48
            Top             =   150
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   5424
            View            =   3
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
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre del Artículo"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ancho"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "alto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "largo"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lv_detalle_agrupadores 
            Height          =   3075
            Index           =   1
            Left            =   45
            TabIndex        =   49
            Top             =   150
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   5424
            View            =   3
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
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre de la Linea"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ancho"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "alto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "largo"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lv_detalle_agrupadores 
            Height          =   3075
            Index           =   2
            Left            =   60
            TabIndex        =   50
            Top             =   150
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   5424
            View            =   3
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
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre de la Sublinea"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ancho"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "alto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "largo"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lv_detalle_agrupadores 
            Height          =   3075
            Index           =   3
            Left            =   45
            TabIndex        =   51
            Top             =   150
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   5424
            View            =   3
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
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre del Tipo de Producto"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ancho"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "alto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "largo"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lv_detalle_agrupadores 
            Height          =   3075
            Index           =   4
            Left            =   45
            TabIndex        =   52
            Top             =   150
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   5424
            View            =   3
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
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre del Subtipo de Producto"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ancho"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "alto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "largo"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de agrupadores"
         Height          =   3270
         Index           =   3
         Left            =   45
         TabIndex        =   36
         Top             =   600
         Width           =   5745
         Begin VB.OptionButton opt_tipoagrupador 
            Caption         =   "Artículos individuales"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   225
            Width           =   2175
         End
         Begin VB.OptionButton opt_tipoagrupador 
            Caption         =   "Linea"
            Height          =   255
            Index           =   1
            Left            =   165
            TabIndex        =   7
            Top             =   885
            Width           =   810
         End
         Begin VB.OptionButton opt_tipoagrupador 
            Caption         =   "Sublinea"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   8
            Top             =   1140
            Width           =   1035
         End
         Begin VB.OptionButton opt_tipoagrupador 
            Caption         =   "Producto"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   9
            Top             =   2010
            Width           =   1065
         End
         Begin VB.OptionButton opt_tipoagrupador 
            Caption         =   "Subtipo de producto"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   10
            Top             =   2265
            Width           =   2175
         End
         Begin VB.ComboBox cmb_detalle_agrupadores 
            Height          =   315
            Index           =   0
            Left            =   1335
            TabIndex        =   11
            Top             =   465
            Width           =   4290
         End
         Begin VB.ComboBox cmb_detalle_agrupadores 
            Height          =   315
            Index           =   1
            Left            =   1260
            TabIndex        =   12
            Top             =   1320
            Width           =   4350
         End
         Begin VB.ComboBox cmb_detalle_agrupadores 
            Height          =   315
            Index           =   2
            Left            =   1260
            TabIndex        =   13
            Top             =   1650
            Width           =   4350
         End
         Begin VB.ComboBox cmb_detalle_agrupadores 
            Height          =   315
            Index           =   3
            Left            =   1305
            TabIndex        =   14
            Top             =   2505
            Width           =   4290
         End
         Begin VB.ComboBox cmb_detalle_agrupadores 
            Height          =   315
            Index           =   4
            Left            =   1665
            TabIndex        =   15
            Top             =   2835
            Width           =   3930
         End
         Begin VB.TextBox txt_detalle_agrupadores 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   41
            Top             =   465
            Width           =   2025
         End
         Begin VB.TextBox txt_detalle_agrupadores 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1275
            MaxLength       =   50
            TabIndex        =   40
            Top             =   1320
            Width           =   2025
         End
         Begin VB.TextBox txt_detalle_agrupadores 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1665
            Width           =   2025
         End
         Begin VB.TextBox txt_detalle_agrupadores 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   38
            Top             =   2520
            Width           =   2025
         End
         Begin VB.TextBox txt_detalle_agrupadores 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   37
            Top             =   2835
            Width           =   2025
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Index           =   7
            Left            =   765
            TabIndex        =   46
            Top             =   510
            Width           =   540
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Linea:"
            Height          =   195
            Index           =   6
            Left            =   795
            TabIndex        =   45
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Sublinea:"
            Height          =   195
            Index           =   5
            Left            =   585
            TabIndex        =   44
            Top             =   1650
            Width           =   660
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Index           =   3
            Left            =   555
            TabIndex        =   43
            Top             =   2550
            Width           =   690
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Subtipo de producto:"
            Height          =   195
            Index           =   4
            Left            =   45
            TabIndex        =   42
            Top             =   2880
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2385
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   35
         Top             =   150
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.Toolbar Toolbar_detalle_agrupadores 
         Height          =   330
         Left            =   60
         TabIndex        =   53
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1(2)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo Registro"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Guardar Registro"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Deshacer cambios"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar Registro"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Index           =   2
         Left            =   30
         TabIndex        =   54
         Top             =   465
         Width           =   11325
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Agrupadores "
      Height          =   3660
      Index           =   0
      Left            =   5520
      TabIndex        =   24
      Top             =   0
      Width           =   6000
      Begin VB.Frame Frame3 
         Height          =   2025
         Index           =   1
         Left            =   45
         TabIndex        =   31
         Top             =   1575
         Width           =   5880
         Begin MSComctlLib.ImageList icono_encabezado 
            Index           =   1
            Left            =   735
            Top             =   -60
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
                  Picture         =   "frmagrupadores2.frx":5B46
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":6420
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lv_agrupadores 
            Height          =   1800
            Left            =   30
            TabIndex        =   5
            Top             =   180
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   3175
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   8184
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "tipo"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Index           =   1
            Left            =   1290
            Top             =   -45
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
                  Picture         =   "frmagrupadores2.frx":6CFA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":75D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":7EAE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":844A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":8D24
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":95FE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":9ED8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":A1F2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":A50C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":AAA8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":ADC2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":AED4
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Agrupadores "
         Height          =   990
         Index           =   2
         Left            =   30
         TabIndex        =   26
         Top             =   585
         Width           =   5910
         Begin VB.TextBox txt_agrupadores 
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   50
            TabIndex        =   3
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox txt_agrupadores 
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   50
            TabIndex        =   4
            Top             =   540
            Width           =   3840
         End
         Begin VB.TextBox txt_agrupadores 
            Height          =   285
            Index           =   2
            Left            =   5325
            MaxLength       =   1
            TabIndex        =   27
            Top             =   540
            Width           =   405
         End
         Begin VB.Label lab_estados 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Index           =   1
            Left            =   465
            TabIndex        =   30
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   29
            Top             =   585
            Width           =   600
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   2
            Left            =   4845
            TabIndex        =   28
            Top             =   585
            Width           =   360
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3375
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   25
         Top             =   30
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.Toolbar Toolbar_agrupadores 
         Height          =   330
         Left            =   60
         TabIndex        =   32
         Top             =   210
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1(1)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo Registro"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Guardar Registro"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Deshacer cambios"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar Registro"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Clonar agrupador"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Index           =   1
         Left            =   15
         TabIndex        =   33
         Top             =   465
         Width           =   5955
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Familia de agrupadores "
      Height          =   3645
      Index           =   0
      Left            =   105
      TabIndex        =   16
      Top             =   0
      Width           =   5310
      Begin VB.Frame Frame3 
         Height          =   2010
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   1575
         Width           =   5115
         Begin MSComctlLib.ImageList ImageList1 
            Index           =   0
            Left            =   600
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   11
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":B46E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":BD48
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":C622
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":CBBE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":D498
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":DD72
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":E64C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":E966
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":EC80
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":F21C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":F536
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList icono_encabezado 
            Index           =   0
            Left            =   0
            Top             =   0
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
                  Picture         =   "frmagrupadores2.frx":F648
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmagrupadores2.frx":FF22
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lv_familia_agrupadores 
            Height          =   1770
            Left            =   30
            TabIndex        =   2
            Top             =   180
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   6967
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Tipo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "familia"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2775
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   17
         Top             =   90
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.Toolbar toolbarfamilia_agrupadores 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   225
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1(0)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo Registro"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Guardar Registro"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Deshacer cambios"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar Registro"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Index           =   0
         Left            =   105
         TabIndex        =   23
         Top             =   480
         Width           =   5145
      End
      Begin VB.Frame Frame1 
         Caption         =   " Familia de agrupadores "
         Height          =   990
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   585
         Width           =   5130
         Begin VB.TextBox txt_familia_agrupadores 
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   50
            TabIndex        =   1
            Top             =   540
            Width           =   4080
         End
         Begin VB.TextBox txt_familia_agrupadores 
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   50
            TabIndex        =   0
            Top             =   225
            Width           =   1200
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   20
            Top             =   540
            Width           =   600
         End
         Begin VB.Label lab_estados 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Index           =   0
            Left            =   465
            TabIndex        =   19
            Top             =   240
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmagrupadores2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios_familia_agrupadores As Boolean
Dim var_hubo_cambios_agrupadores As Boolean
Dim var_hubo_cambios As Boolean
Dim vartipoagrupador As Integer
Dim var_guardar_cambios_familia_agrupadores As Boolean
Dim var_guardar_cambios_agrupadores As Boolean
Dim var_guardar_cambios_detalle_agrupadores As Boolean

Private Sub Combo1_Click()
   txt_familia_agrupadores(0) = Obtener_llave(cnn, rsaux, "TB_PAISES", "VCHA_PAI_NOMBRE", Combo1, 0, "T")
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
  
   If KeyAscii = 27 Then
      End
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
    var_guardar_cambios_detalle_agrupadores = False
    var_guardar_cambios_agrupadores = False
    var_guardar_cambios_familia_agrupadores = False
    var_modifica_registro_familia_agrupadores = True
    lv_familia_agrupadores.SmallIcons = ImageList1(0)
   ' Call pro_encabezadosView(Me, lv_familia_agrupadores, False)
    Call pro_llena_listview1_familia_agrupadores
    Call pro_textos_familia_agrupadores

'    Call pro_AsignarAViewColor(lv_familia_agrupadores, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
    If rs.BOF Then
       toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = False
       toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = False
       toolbarfamilia_agrupadores.Buttons.Item(4).Enabled = False
    Else
       toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = True
       toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = True
       toolbarfamilia_agrupadores.Buttons.Item(4).Enabled = True
    End If
    rs.Close
   varfamiliaagrupadores = frmfamilia_agrupadores.txt_familia_agrupadores(0)
   var_modifica_registro_agrupadores = True
   Call pro_llena_listview1_agrupadores
   pro_textos_agrupadores
   rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      Toolbar_agrupadores.Buttons.Item(2).Enabled = False
      Toolbar_agrupadores.Buttons.Item(3).Enabled = False
      Toolbar_agrupadores.Buttons.Item(4).Enabled = False
   Else
      Toolbar_agrupadores.Buttons.Item(2).Enabled = True
      Toolbar_agrupadores.Buttons.Item(3).Enabled = True
      Toolbar_agrupadores.Buttons.Item(4).Enabled = True
   End If
   rs.Close
   varagrupador = frmagrupadores.txt_agrupadores(0)
   var_modifica_registro = True
   rs.Open "select * from tb_articulos", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_detalle_agrupadores(0).hwnd, rs, 1)
   rs.Close
   rs.Open "select * from tb_lineas", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_detalle_agrupadores(1).hwnd, rs, 1)
   rs.Close
   rs.Open "select * from tb_sublineas", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_detalle_agrupadores(2).hwnd, rs, 2)
   rs.Close
   rs.Open "select * from tb_productos", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_detalle_agrupadores(3).hwnd, rs, 1)
   rs.Close
   rs.Open "select * from tb_tipoarticulos", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_detalle_agrupadores(4).hwnd, rs, 1)
   rs.Close
   rs.Open "select * from TB_DETALLE_AGRUPADORES where vcha_agr_agrupador_id = '" & varagrupador & "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      vartipoagrupador = rs(1).Value
      If vartipoagrupador = 1 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(0), False)
         cmb_detalle_agrupadores(0).Enabled = True
         cmb_detalle_agrupadores(1).Enabled = False
         cmb_detalle_agrupadores(2).Enabled = False
         cmb_detalle_agrupadores(3).Enabled = False
         cmb_detalle_agrupadores(4).Enabled = False
         lv_detalle_agrupadores(0).Visible = True
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_art_articulo_id,b.vcha_art_nombre_español from TB_DETALLE_AGRUPADORES a, TB_ARTICULOS b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_art_articulo_id = b.vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = True
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 2 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(1), False)
         cmb_detalle_agrupadores(0).Enabled = False
         cmb_detalle_agrupadores(1).Enabled = True
         cmb_detalle_agrupadores(2).Enabled = False
         cmb_detalle_agrupadores(3).Enabled = False
         cmb_detalle_agrupadores(4).Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = True
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_lin_linea_id,b.vcha_lin_nombre from TB_DETALLE_AGRUPADORES a, TB_lineas b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_lin_linea_id = b.vcha_lin_linea_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = True
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 3 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(2), False)
         cmb_detalle_agrupadores(0).Enabled = False
         cmb_detalle_agrupadores(1).Enabled = True
         cmb_detalle_agrupadores(2).Enabled = True
         cmb_detalle_agrupadores(3).Enabled = False
         cmb_detalle_agrupadores(4).Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = True
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_sli_sublinea_id,b.vcha_sli_nombre from TB_DETALLE_AGRUPADORES a, TB_sublineas b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_sli_sublinea_id = b.vcha_sli_sublinea_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = True
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 4 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(3), False)
         cmb_detalle_agrupadores(0).Enabled = False
         cmb_detalle_agrupadores(1).Enabled = False
         cmb_detalle_agrupadores(2).Enabled = False
         cmb_detalle_agrupadores(3).Enabled = True
         cmb_detalle_agrupadores(4).Enabled = False
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = True
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_pro_producto_id,b.vcha_pro_nombre from TB_DETALLE_AGRUPADORES a, TB_productos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_pro_producto_id = b.vcha_pro_producto_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = True
         opt_tipoagrupador(4).Value = False
      End If
      If vartipoagrupador = 5 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(4), False)
         cmb_detalle_agrupadores(0).Enabled = False
         cmb_detalle_agrupadores(1).Enabled = False
         cmb_detalle_agrupadores(2).Enabled = False
         cmb_detalle_agrupadores(3).Enabled = True
         cmb_detalle_agrupadores(4).Enabled = True
         lv_detalle_agrupadores(0).Visible = False
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = True
         rs.Open "select a.vcha_tar_tipo_articulo_id,b.vcha_tar_nombre from TB_DETALLE_AGRUPADORES a, TB_tipoarticulos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_tar_tipo_articulo_id = b.vcha_tar_tipo_articulo_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = False
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = True
      End If
   Else
      vartipoagrupador = 1
      If vartipoagrupador = 1 Then
         rs.Close
         Call pro_encabezadosView(Me, lv_detalle_agrupadores(0), False)
         Call pro_llena_listview1
         pro_textos
         cmb_detalle_agrupadores(0).Enabled = True
         cmb_detalle_agrupadores(1).Enabled = False
         cmb_detalle_agrupadores(2).Enabled = False
         cmb_detalle_agrupadores(3).Enabled = False
         cmb_detalle_agrupadores(4).Enabled = False
         lv_detalle_agrupadores(0).Visible = True
         lv_detalle_agrupadores(1).Visible = False
         lv_detalle_agrupadores(2).Visible = False
         lv_detalle_agrupadores(3).Visible = False
         lv_detalle_agrupadores(4).Visible = False
         rs.Open "select a.vcha_art_articulo_id,b.vcha_art_nombre_español from TB_DETALLE_AGRUPADORES a, TB_ARTICULOS b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_art_articulo_id = b.vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
         opt_tipoagrupador(0).Value = True
         opt_tipoagrupador(1).Value = False
         opt_tipoagrupador(2).Value = False
         opt_tipoagrupador(3).Value = False
         opt_tipoagrupador(4).Value = False
      End If
   End If
   If rs.BOF Then
      Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = False
      Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = False
      Toolbar_detalle_agrupadores.Buttons.Item(4).Enabled = False
    Else
      Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = True
      Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = True
      Toolbar_detalle_agrupadores.Buttons.Item(4).Enabled = True
    End If
    rs.Close
    Call pro_llena_listview1
    pro_textos

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim var_salir As Boolean
Dim var_familia As Integer
Dim var_agrupador As Integer
Dim var_detalle As Integer
    If var_guardar_cambios_familia_agrupadores = True Then
       var_familia = MsgBox("¿Deseas guardar los cambios de la famila de agrupadores?", vbOKCancel, "ATENCION")
       If var_familia = 1 Then
          Call pro_guardar_familia_agrupadores
       End If
    End If
    If var_guardar_cambios_agrupadores = True Then
       var_agrupador = MsgBox("¿Deseas guardar los cambios del agrupador?", vbOKCancel, "ATENCION")
       If var_agrupador = 1 Then
          Call pro_guardar_agrupadores
       End If
    End If
    If var_guardar_cambios_detalle_agrupadores = True Then
       var_detalle = MsgBox("¿Deseas guardar los cambios del detalle del agrupador?", vbOKCancel, "ATENCION")
       If var_detalle = 1 Then
          Call pro_guardar_detalle_agrupadores
       End If
    End If
    var_swpassword = False
    var_modifica_registro_familia_agrupadores = False
End Sub

Private Sub lv_familia_agrupadores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_familia_agrupadores.selectedItem = Item
        pro_textos_familia_agrupadores
        var_modifica_registro_familia_agrupadores = True
        txt_familia_agrupadores(0).Enabled = False
        lv_agrupadores.ListItems.Clear
        Call pro_llena_listview1_agrupadores
        lv_detalle_agrupadores(0).ListItems.Clear
        lv_detalle_agrupadores(1).ListItems.Clear
        lv_detalle_agrupadores(2).ListItems.Clear
        lv_detalle_agrupadores(3).ListItems.Clear
        lv_detalle_agrupadores(4).ListItems.Clear
        txt_agrupadores(0) = ""
        txt_agrupadores(1) = ""
        pro_textos_agrupadores
        If txt_agrupadores(0) <> "" Then
           cmb_detalle_agrupadores(0).Text = ""
           cmb_detalle_agrupadores(1).Text = ""
           cmb_detalle_agrupadores(2).Text = ""
           cmb_detalle_agrupadores(3).Text = ""
           cmb_detalle_agrupadores(4).Text = ""
           txt_detalle_agrupadores(0) = ""
           txt_detalle_agrupadores(1) = ""
           txt_detalle_agrupadores(2) = ""
           txt_detalle_agrupadores(3) = ""
           txt_detalle_agrupadores(4) = ""
          Call pro_llena_listview1
        End If
          
End Sub

Private Sub tool_atras_siguiente_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
   Dim varagrupador As String
   If Index = 1 Then
      frmagrupadores.Caption = "Agrupadores de artículos de la familia de:  " + lv_familia_agrupadores.selectedItem.SubItems(1)
      frmagrupadores.Show 1
   End If
   
   If Index = 0 Then
      Select Case Button.Index
      Case 1
          lv_familia_agrupadores.SetFocus
          Call pro_avanzar(Me, lv_familia_agrupadores, Button)
          pro_textos_familia_agrupadores
      Case 2
          lv_familia_agrupadores.SetFocus
          Call pro_avanzar(Me, lv_familia_agrupadores, Button)
          pro_textos_familia_agrupadores
      Case 3
          Call pro_busca_registro(lv_familia_agrupadores, txt_buscar, False)
          txt_buscar = ""
          pro_textos_familia_agrupadores
      End Select
    End If
End Sub

Private Sub toolbarfamilia_agrupadores_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_familia_agrupadores(0).Enabled = True
        txt_familia_agrupadores(0).SetFocus: var_modifica_registro_familia_agrupadores = False
        toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = True
        toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = True
        var_guardar_cambios_familia_agrupadores = True
    Case 2
        If txt_familia_agrupadores(0) = "" Or txt_familia_agrupadores(1) = "" Then
           MsgBox "Falta información", vbOKOnly, "ATENCION"
        Else
           var_resultado = InStr(1, var_menus, Me.Caption)
           var_inicio = var_resultado + Len(Me.Caption) + 3
           If Mid(var_menus, var_inicio, 1) = "1" Then
              Set var_forma = frmfamilia_agrupadores
              var_swpassword = True
              sw_primera_validacion = False
              frmpasswords.Show 1
           Else
              If Mid(var_menus, var_inicio, 2) = "01" Then
                 Set var_forma = frmfamilia_agrupadores
                 var_swpassword = True
                 sw_primera_validacion = False
                 frmpasswords2.txt_supervisor = var_supervisor
                 frmpasswords2.Show 1
              Else
                 Call pro_guardar_familia_agrupadores
                 rs.Open "select * from tb_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
                 If rs.BOF Then
                    toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = False
                    toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = False
                    toolbarfamilia_agrupadores.Buttons.Item(4).Enabled = False
                    var_guardar_cambios_familia_agrupadores = False
                 Else
                    toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = True
                    toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = True
                    toolbarfamilia_agrupadores.Buttons.Item(4).Enabled = True
                    var_guardar_cambios_familia_agrupadores = False
                 End If
                 rs.Close
              End If
            End If
         End If
    Case 3
       Call pro_textos_familia_agrupadores
    Case 4
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmfamilia_agrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmfamilia_agrupadores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_familia_agrupadores
               rs.Open "select * from tb_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
               If rs.BOF Then
                  toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = False
                  toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = False
                  toolbarfamilia_agrupadores.Buttons.Item(4).Enabled = False
               Else
                  toolbarfamilia_agrupadores.Buttons.Item(2).Enabled = True
                  toolbarfamilia_agrupadores.Buttons.Item(3).Enabled = True
                  toolbarfamilia_agrupadores.Buttons.Item(4).Enabled = True
               End If
               rs.Close
            End If
        End If
    Case 6
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_familia_agrupadores, "LISTADO DE familia_agrupadores")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_familia_agrupadores()

Dim ok As Boolean

Set TB_FAMILIA_AGRUPADORES = New TB_FAMILIA_AGRUPADORES
    
    ok = True
    If txt_familia_agrupadores(0) <> "" And txt_familia_agrupadores(1) <> "" Then
        If var_hubo_cambios_familia_agrupadores Then
            ok = TB_FAMILIA_AGRUPADORES.Anadir(txt_familia_agrupadores(0), txt_familia_agrupadores(1))
            If ok Then
                pro_actualiza_ListView_familia_agrupadores
                txt_familia_agrupadores(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_familia_agrupadores.ListItems.Count
                var_modifica_registro_familia_agrupadores = True
            Else
                MsgBox "No se puede grabar registro: " + TB_FAMILIA_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_FAMILIA_AGRUPADORES = Nothing: var_hubo_cambios_familia_agrupadores = False

End Sub

Sub pro_elimina_familia_agrupadores()
   Dim var_llave_usuarios As String
   Set TB_FAMILIA_AGRUPADORES = New TB_FAMILIA_AGRUPADORES
   ok = True
   On Error GoTo salir:
   If txt_familia_agrupadores(0) <> "" And txt_familia_agrupadores(1) <> "" And var_modifica_registro_familia_agrupadores = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_FAMILIA_AGRUPADORES.Eliminar(txt_familia_agrupadores(0))
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_familia_agrupadores.ListItems.Remove (lv_familia_agrupadores.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_familia_agrupadores.ListItems.Count
        lv_familia_agrupadores.selectedItem.Selected = True
        pro_textos_familia_agrupadores
      Else
        MsgBox "No se puede eliminar registro: " + TB_FAMILIA_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_FAMILIA_AGRUPADORES = Nothing
End Sub


Sub pro_llena_listview1_familia_agrupadores()
   Dim list_item As ListItem
   rs.Open "select * from TB_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_familia_agrupadores.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
    Wend
    rs.Close
End Sub


Sub pro_textos_familia_agrupadores()
On Error GoTo err0:
        txt_familia_agrupadores(0) = lv_familia_agrupadores.selectedItem
        txt_familia_agrupadores(1) = lv_familia_agrupadores.selectedItem.SubItems(1)
        
err0:
End Sub

Private Sub pro_actualiza_ListView_familia_agrupadores()
Dim list_item As ListItem

    If var_modifica_registro_familia_agrupadores = False Then
        Set list_item = lv_familia_agrupadores.ListItems.Add(, , txt_familia_agrupadores(0))
        list_item.SubItems(1) = txt_familia_agrupadores(1)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_familia_agrupadores.ListItems.Item(lv_familia_agrupadores.selectedItem.Index).Checked = False
        lv_familia_agrupadores.ListItems.Item(lv_familia_agrupadores.selectedItem.Index) = txt_familia_agrupadores(0)
        lv_familia_agrupadores.ListItems.Item(lv_familia_agrupadores.selectedItem.Index).ListSubItems(1) = txt_familia_agrupadores(1)
        lv_familia_agrupadores.ListItems.Item(lv_familia_agrupadores.selectedItem.Index).Selected = True
    End If
'    lv_familia_agrupadores.SetFocus
End Sub



Private Sub txt_familia_agrupadores_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios_familia_agrupadores = True
   var_guardar_cambios_familia_agrupadores = True
End Sub






Private Sub lv_agrupadores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_agrupadores.selectedItem = Item
        pro_textos_agrupadores
        var_modifica_registro_agrupadores = True
        txt_agrupadores(0).Enabled = False
        lv_detalle_agrupadores(0).ListItems.Clear
        lv_detalle_agrupadores(1).ListItems.Clear
        lv_detalle_agrupadores(2).ListItems.Clear
        lv_detalle_agrupadores(3).ListItems.Clear
        lv_detalle_agrupadores(4).ListItems.Clear
        cmb_detalle_agrupadores(0).Text = ""
        cmb_detalle_agrupadores(1).Text = ""
        cmb_detalle_agrupadores(2).Text = ""
        cmb_detalle_agrupadores(3).Text = ""
        cmb_detalle_agrupadores(4).Text = ""
        txt_detalle_agrupadores(0) = ""
        txt_detalle_agrupadores(1) = ""
        txt_detalle_agrupadores(2) = ""
        txt_detalle_agrupadores(3) = ""
        txt_detalle_agrupadores(4) = ""
        Call pro_llena_listview1
End Sub


Private Sub Toolbar_agrupadores_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
   Case 1
      Call pro_limpiatextos(Me)
      txt_agrupadores(0).Enabled = True
      txt_agrupadores(0).SetFocus: var_modifica_registro_agrupadores = False
      Toolbar_agrupadores.Buttons.Item(2).Enabled = True
      Toolbar_agrupadores.Buttons.Item(3).Enabled = True
      var_guardar_cambios_agrupadores = False
   Case 2
      If txt_agrupadores(0) = "" Or txt_agrupadores(1) = "" Or txt_agrupadores(2) = "" Then
         MsgBox "Información incompleta", vbOKOnly, "ATENCION"
      Else
      var_resultado = InStr(1, var_menus, Me.Caption)
      var_inicio = var_resultado + Len(Me.Caption) + 3
      If Mid(var_menus, var_inicio, 1) = "1" Then
         Set var_forma = frmagrupadores
         var_swpassword = True
         sw_primera_validacion = False
         frmpasswords.Show 1
      Else
         If Mid(var_menus, var_inicio, 2) = "01" Then
            Set var_forma = frmagrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords2.txt_supervisor = var_supervisor
            frmpasswords2.Show 1
         Else
            Call pro_guardar_agrupadores
            rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.BOF Then
               Toolbar_agrupadores.Buttons.Item(2).Enabled = False
               Toolbar_agrupadores.Buttons.Item(3).Enabled = False
               Toolbar_agrupadores.Buttons.Item(4).Enabled = False
               var_guardar_cambios_agrupadores = True
            Else
               Toolbar_agrupadores.Buttons.Item(2).Enabled = True
               Toolbar_agrupadores.Buttons.Item(3).Enabled = True
               Toolbar_agrupadores.Buttons.Item(4).Enabled = True
               var_guardar_cambios_agrupadores = True
            End If
            rs.Close
            lv_agrupadores.ListItems.Clear
            Call pro_llena_listview1_agrupadores
         End If
      End If
      End If
   Case 3
      Call pro_textos_agrupadores
   Case 4
      var_resultado = InStr(1, var_menus, Me.Caption)
      var_inicio = var_resultado + Len(Me.Caption) + 3
      If Mid(var_menus, var_inicio, 1) = "1" Then
         Set var_forma = frmagrupadores
         var_swpassword = True
         sw_primera_validacion = False
         frmpasswords.Show 1
      Else
         If Mid(var_menus, var_inicio, 2) = "01" Then
            Set var_forma = frmagrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords2.txt_supervisor = var_supervisor
            frmpasswords2.Show 1
         Else
            Call pro_elimina_agrupadores
            rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.BOF Then
               Toolbar_agrupadores.Buttons.Item(2).Enabled = False
               Toolbar_agrupadores.Buttons.Item(3).Enabled = False
               Toolbar_agrupadores.Buttons.Item(4).Enabled = False
            Else
               Toolbar_agrupadores.Buttons.Item(2).Enabled = True
               Toolbar_agrupadores.Buttons.Item(3).Enabled = True
               Toolbar_agrupadores.Buttons.Item(4).Enabled = True
            End If
            rs.Close
         End If
      End If
   Case 5
      frmclonacionagrupadores.Show 1
      lv_agrupadores.ListItems.Clear
      Call pro_llena_listview1_agrupadores
   Case 7
      If vector_valida_passwords(var_indice_menu) = "*" Then
         frmpasswords.Show
      Else
         Call gPrintListView(lv_agrupadores, "LISTADO DE agrupadores")
      End If
   Case 9
      Unload Me
   End Select
End Sub

Sub pro_guardar_agrupadores()

Dim ok As Boolean

Set TB_AGRUPADORES = New TB_AGRUPADORES
    
    ok = True
    If txt_agrupadores(0) <> "" And txt_agrupadores(1) <> "" And txt_agrupadores(2) <> "" Then
        If var_hubo_cambios_agrupadores Then
            varfamiliaagrupadores = lv_familia_agrupadores.selectedItem
            ok = TB_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), txt_agrupadores(1), txt_agrupadores(2))
            If ok Then
                pro_actualiza_ListView
                txt_agrupadores(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_agrupadores.ListItems.Count
                var_modifica_registro_agrupadores = True
            Else
                MsgBox "No se puede grabar registro: " + TB_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_AGRUPADORES = Nothing: var_hubo_cambios_agrupadores = False

End Sub

Sub pro_elimina_agrupadores()
   Dim var_llave_usuarios As String
   Set TB_AGRUPADORES = New TB_AGRUPADORES
   ok = True
   On Error GoTo salir:
   
   If txt_agrupadores(0) <> "" And txt_agrupadores(1) <> "" And txt_agrupadores(2) _
      <> "" And txt_agrupadores(2) <> "" And var_modifica_registro_agrupadores = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_AGRUPADORES.Eliminar(txt_agrupadores(0))
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_agrupadores.ListItems.Remove (lv_agrupadores.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_agrupadores.ListItems.Count
        lv_agrupadores.selectedItem.Selected = True
        pro_textos_agrupadores
      Else
        MsgBox "No se puede eliminar registro: " + TB_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_AGRUPADORES = Nothing
End Sub


Sub pro_llena_listview1_agrupadores()
   Dim list_item As ListItem
   rs.Open "select * from TB_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      varfamiliaagrupadores = lv_familia_agrupadores.selectedItem
   Else
      varfamiliaagrupadores = 0
   End If
   rs.Close
   rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_agrupadores.ListItems.Add(, , rs(1).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
      rs.MoveNext:
    Wend
    rs.Close
    pro_textos
End Sub


Sub pro_textos_agrupadores()
On Error GoTo err0:
        txt_agrupadores(0) = lv_agrupadores.selectedItem
        txt_agrupadores(1) = lv_agrupadores.selectedItem.SubItems(1)
        txt_agrupadores(2) = lv_agrupadores.selectedItem.SubItems(2)
err0:
End Sub

Private Sub pro_actualiza_ListView_agrupadores()
Dim list_item As ListItem
    If var_modifica_registro_agrupadores = False Then
        Set list_item = lv_agrupadores.ListItems.Add(, , txt_agrupadores(0))
        list_item.SubItems(1) = txt_agrupadores(1)
        list_item.SubItems(2) = txt_agrupadores(2)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).Checked = False
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index) = txt_agrupadores(0)
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(1) = txt_agrupadores(1)
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(2) = txt_agrupadores(2)
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).Selected = True
    End If
End Sub

Private Sub txt_agrupadores_Change(Index As Integer)
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios_agrupadores = True
End Sub

Private Sub txt_agrupadores_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios_agrupadores = True
   var_guardar_cambios_agrupadores = True
End Sub




Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_detalle_agrupadores, txt_buscar, False)
    txt_buscar = ""
    pro_textos
End Sub





Private Sub cmb_detalle_agrupadores_Click(Index As Integer)
   var_guardar_cambios_detalle_agrupadores = True
   If Index = 0 Then
      txt_detalle_agrupadores(0) = Obtener_llave(cnn, rs, "TB_articulos", "VCHA_ART_NOMBRE_ESPAÑOL", cmb_detalle_agrupadores(0), 0, "T")
      vardetallearticulo = cmb_detalle_agrupadores(0).Text
   End If
   If Index = 1 Then
      txt_detalle_agrupadores(1) = Obtener_llave(cnn, rs, "TB_LINEAS", "VCHA_LIN_nombre", cmb_detalle_agrupadores(1), 0, "T")
      vardetallelinea = cmb_detalle_agrupadores(1).Text
   End If
   If Index = 2 Then
      txt_detalle_agrupadores(2) = Obtener_llave(cnn, rs, "TB_SUBLINEAS", "VCHA_SLI_NOMBRE", cmb_detalle_agrupadores(2), 1, "T")
      vardetallelinea = cmb_detalle_agrupadores(1).Text
      vardetallesublinea = cmb_detalle_agrupadores(2).Text
   End If
   If Index = 3 Then
      txt_detalle_agrupadores(3) = Obtener_llave(cnn, rs, "TB_PRODUCTOS", "VCHA_PRO_NOMBRE", cmb_detalle_agrupadores(3), 0, "T")
   End If
   If Index = 4 Then
      txt_detalle_agrupadores(4) = Obtener_llave(cnn, rs, "TB_TIPOARTICULOS", "VCHA_TAR_NOMBRE", cmb_detalle_agrupadores(4), 0, "T")
   End If
End Sub



Private Sub lv_detalle_agrupadores_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
   If opt_tipoagrupador(0).Value = True Then
      Set lv_detalle_agrupadores(0).selectedItem = Item
      var_modifica_registro = True
      txt_detalle_agrupadores(0).Enabled = False
   End If
   If opt_tipoagrupador(1).Value = True Then
      Set lv_detalle_agrupadores(1).selectedItem = Item
      var_modifica_registro = True
      txt_detalle_agrupadores(1).Enabled = False
   End If
   If opt_tipoagrupador(2).Value = True Then
      Set lv_detalle_agrupadores(2).selectedItem = Item
      var_modifica_registro = True
      txt_detalle_agrupadores(2).Enabled = False
   End If
   If opt_tipoagrupador(3).Value = True Then
      Set lv_detalle_agrupadores(3).selectedItem = Item
      var_modifica_registro = True
      txt_detalle_agrupadores(3).Enabled = False
   End If
   If opt_tipoagrupador(4).Value = True Then
      Set lv_detalle_agrupadores(4).selectedItem = Item
      var_modifica_registro = True
      txt_detalle_agrupadores(4).Enabled = False
   End If
   pro_textos
End Sub


Private Sub opt_tipoagrupador_Click(Index As Integer)
   If opt_tipoagrupador(0).Value = True Then
      cmb_detalle_agrupadores(0).Enabled = True
      cmb_detalle_agrupadores(1).Enabled = False
      cmb_detalle_agrupadores(2).Enabled = False
      cmb_detalle_agrupadores(3).Enabled = False
      cmb_detalle_agrupadores(4).Enabled = False
      lv_detalle_agrupadores(0).Visible = True
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = False
      pro_textos
   End If
   If opt_tipoagrupador(1).Value = True Then
      cmb_detalle_agrupadores(0).Enabled = False
      cmb_detalle_agrupadores(1).Enabled = True
      cmb_detalle_agrupadores(2).Enabled = False
      cmb_detalle_agrupadores(3).Enabled = False
      cmb_detalle_agrupadores(4).Enabled = False
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = True
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = False
      pro_textos
   End If
   If opt_tipoagrupador(2).Value = True Then
      cmb_detalle_agrupadores(0).Enabled = False
      cmb_detalle_agrupadores(1).Enabled = True
      cmb_detalle_agrupadores(2).Enabled = True
      cmb_detalle_agrupadores(3).Enabled = False
      cmb_detalle_agrupadores(4).Enabled = False
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = True
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = False
      pro_textos
   End If
   If opt_tipoagrupador(3).Value = True Then
      cmb_detalle_agrupadores(0).Enabled = False
      cmb_detalle_agrupadores(1).Enabled = False
      cmb_detalle_agrupadores(2).Enabled = False
      cmb_detalle_agrupadores(3).Enabled = True
      cmb_detalle_agrupadores(4).Enabled = False
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = True
      lv_detalle_agrupadores(4).Visible = False
      pro_textos
   End If
   If opt_tipoagrupador(4).Value = True Then
      cmb_detalle_agrupadores(0).Enabled = False
      cmb_detalle_agrupadores(1).Enabled = False
      cmb_detalle_agrupadores(2).Enabled = False
      cmb_detalle_agrupadores(3).Enabled = True
      cmb_detalle_agrupadores(4).Enabled = True
      lv_detalle_agrupadores(0).Visible = False
      lv_detalle_agrupadores(1).Visible = False
      lv_detalle_agrupadores(2).Visible = False
      lv_detalle_agrupadores(3).Visible = False
      lv_detalle_agrupadores(4).Visible = True
      pro_textos
   End If
End Sub

Private Sub Toolbar_detalle_agrupadores_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_detalle_agrupadores(0).Enabled = True
        txt_detalle_agrupadores(0).SetFocus: var_modifica_registro = False
        Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = True
        Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = True
        var_guardar_cambios_detalle_agrupadores = True
    Case 2
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmdetalle_agrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmdetalle_agrupadores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_guardar_detalle_agrupadores
               rs.Open "select * from tb_detalle_agrupadores", cnn, adOpenDynamic, adLockOptimistic
               If rs.BOF Then
                  Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = False
                  Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = False
                  Toolbar_detalle_agrupadores.Buttons.Item(4).Enabled = False
                  var_guardar_cambios_detalle_agrupadores = False
               Else
                  Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = True
                  Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = True
                  Toolbar_detalle_agrupadores.Buttons.Item(4).Enabled = True
                  var_guardar_cambios_detalle_agrupadores = False
               End If
               rs.Close
            End If
        End If
    Case 3
       Call pro_textos
    Case 4
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmdetalle_agrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmdetalle_agrupadores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_detalle_agrupadores
               rs.Open "select * from tb_detalle_agrupadores ", cnn, adOpenDynamic, adLockOptimistic
               If rs.BOF Then
                  Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = False
                  Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = False
                  Toolbar_detalle_agrupadores.Buttons.Item(4).Enabled = False
               Else
                  Toolbar_detalle_agrupadores.Buttons.Item(2).Enabled = True
                  Toolbar_detalle_agrupadores.Buttons.Item(3).Enabled = True
                  Toolbar_detalle_agrupadores.Buttons.Item(4).Enabled = True
               End If
               rs.Close
            End If
        End If
    Case 6
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_detalle_agrupadores, "LISTADO DE detalle_agrupadores")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_detalle_agrupadores()
   Dim ok As Boolean
   If opt_tipoagrupador(0).Value = True Then
      vartipoagrupador = 1
   End If
   If opt_tipoagrupador(1).Value = True Then
      vartipoagrupador = 2
   End If
   If opt_tipoagrupador(2).Value = True Then
      vartipoagrupador = 3
   End If
   If opt_tipoagrupador(3).Value = True Then
      vartipoagrupador = 4
   End If
   If opt_tipoagrupador(4).Value = True Then
      vartipoagrupador = 5
   End If
   If vartipoagrupador = 1 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_detalle_agrupadores(0) <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, txt_detalle_agrupadores(0), " ", " ", " ", " ")
            If ok Then
               pro_actualiza_ListView
               txt_detalle_agrupadores(0).Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(0).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 2 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_detalle_agrupadores(1) <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", txt_detalle_agrupadores(1).Text, " ", " ", " ")
            If ok Then
               pro_actualiza_ListView
               txt_detalle_agrupadores(1).Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(1).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 3 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_detalle_agrupadores(1) <> "" And txt_detalle_agrupadores(2) <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", txt_detalle_agrupadores(1), txt_detalle_agrupadores(2), " ", " ")
            If ok Then
               pro_actualiza_ListView
               txt_detalle_agrupadores(2).Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(2).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 4 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_detalle_agrupadores(3) <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", " ", txt_detalle_agrupadores(3), " ")
            If ok Then
               pro_actualiza_ListView
               txt_detalle_agrupadores(3).Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(3).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
   If vartipoagrupador = 5 Then
      Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
      ok = True
      If txt_detalle_agrupadores(3) <> "" And txt_detalle_agrupadores(4) <> "" Then
         If var_hubo_cambios Then
            ok = TB_DETALLE_AGRUPADORES.Anadir(varagrupador, vartipoagrupador, " ", " ", " ", txt_detalle_agrupadores(3), txt_detalle_agrupadores(4))
            If ok Then
               pro_actualiza_ListView
               txt_detalle_agrupadores(4).Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_detalle_agrupadores(4).ListItems.Count
               var_modifica_registro = True
            Else
               MsgBox "No se puede grabar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
         End If
      End If
   End If
Set TB_DETALLE_AGRUPADORES = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_detalle_agrupadores()
   Dim var_llave_usuarios As String
   Set TB_DETALLE_AGRUPADORES = New TB_DETALLE_AGRUPADORES
   On Error GoTo salir:
   ok = True
   If var_modifica_registro = True Then
      If opt_tipoagrupador(0).Value = True Then
         If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = TB_DETALLE_AGRUPADORES.Eliminar(varagrupador, 1, txt_detalle_agrupadores(0), "", "", "", "")
         Else
            GoTo salir:
         End If
         If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            lv_detalle_agrupadores(0).ListItems.Remove (lv_detalle_agrupadores(0).selectedItem.Index)
            Call pro_limpiatextos(Me)
            txt_registros = lv_detalle_agrupadores(0).ListItems.Count
            lv_detalle_agrupadores(0).selectedItem.Selected = True
            pro_textos
          Else
            MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
          End If
      End If
      
      If opt_tipoagrupador(1).Value = True Then
         If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = TB_DETALLE_AGRUPADORES.Eliminar(varagrupador, 2, "", txt_detalle_agrupadores(1), "", "", "")
         Else
            GoTo salir:
         End If
         If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            lv_detalle_agrupadores(1).ListItems.Remove (lv_detalle_agrupadores(1).selectedItem.Index)
            Call pro_limpiatextos(Me)
            txt_registros = lv_detalle_agrupadores(1).ListItems.Count
            lv_detalle_agrupadores(1).selectedItem.Selected = True
            pro_textos
          Else
            MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
          End If
      End If
      
      If opt_tipoagrupador(2).Value = True Then
         If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = TB_DETALLE_AGRUPADORES.Eliminar(varagrupador, 3, "", txt_detalle_agrupadores(1), txt_detalle_agrupadores(2), "", "")
         Else
            GoTo salir:
         End If
         If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            lv_detalle_agrupadores(2).ListItems.Remove (lv_detalle_agrupadores(2).selectedItem.Index)
            Call pro_limpiatextos(Me)
            txt_registros = lv_detalle_agrupadores(2).ListItems.Count
            lv_detalle_agrupadores(2).selectedItem.Selected = True
            pro_textos
          Else
            MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
          End If
      End If
      
      If opt_tipoagrupador(3).Value = True Then
         If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = TB_DETALLE_AGRUPADORES.Eliminar(varagrupador, 4, "", "", "", txt_detalle_agrupadores(3), "")
         Else
            GoTo salir:
         End If
         If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            lv_detalle_agrupadores(3).ListItems.Remove (lv_detalle_agrupadores(3).selectedItem.Index)
            Call pro_limpiatextos(Me)
            txt_registros = lv_detalle_agrupadores(3).ListItems.Count
            lv_detalle_agrupadores(3).selectedItem.Selected = True
            pro_textos
          Else
            MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
          End If
      End If
      
      If opt_tipoagrupador(4).Value = True Then
         If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = TB_DETALLE_AGRUPADORES.Eliminar(varagrupador, 5, "", "", "", txt_detalle_agrupadores(3), txt_detalle_agrupadores(4))
         Else
            GoTo salir:
         End If
         If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            lv_detalle_agrupadores(4).ListItems.Remove (lv_detalle_agrupadores(4).selectedItem.Index)
            Call pro_limpiatextos(Me)
            txt_registros = lv_detalle_agrupadores(4).ListItems.Count
            lv_detalle_agrupadores(4).selectedItem.Selected = True
            pro_textos
          Else
            MsgBox "No se puede eliminar registro: " + TB_DETALLE_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
          End If
      End If
   
   End If
salir:
   Set TB_DETALLE_AGRUPADORES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rsaux2.Open "select distinct inte_dea_tipo from tb_detalle_agrupadores ", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
   varagrupador = lv_agrupadores.selectedItem
   Else
   varagrupador = 0
   End If
   rsaux2.Close
   rsaux2.Open "select distinct inte_dea_tipo from tb_detalle_agrupadores where VCHA_AGR_AGRUPADOR_ID = '" + varagrupador + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
      While Not rsaux2.EOF
         vartipoagrupador = rsaux2(0).Value
         If vartipoagrupador = 1 Then
            lv_detalle_agrupadores(0).Visible = True
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select a.vcha_art_articulo_id,b.vcha_art_nombre_español from TB_DETALLE_AGRUPADORES a, TB_ARTICULOS b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_art_articulo_id = b.vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(0).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            cmb_detalle_agrupadores(0).Enabled = True
            cmb_detalle_agrupadores(1).Enabled = False
            cmb_detalle_agrupadores(2).Enabled = False
            cmb_detalle_agrupadores(3).Enabled = False
            cmb_detalle_agrupadores(4).Enabled = False
            lv_detalle_agrupadores(0).Visible = True
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = True
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 2 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = True
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select distinct a.vcha_lin_linea_id,b.vcha_lin_nombre from TB_DETALLE_AGRUPADORES a, TB_lineas b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_lin_linea_id = b.vcha_lin_linea_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(1).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            cmb_detalle_agrupadores(0).Enabled = False
            cmb_detalle_agrupadores(1).Enabled = True
            cmb_detalle_agrupadores(2).Enabled = False
            cmb_detalle_agrupadores(3).Enabled = False
            cmb_detalle_agrupadores(4).Enabled = False
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = True
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = True
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 3 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = True
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select distinct a.vcha_sli_sublinea_id,b.vcha_sli_nombre,a.vcha_lin_linea_id,c.vcha_lin_nombre from TB_DETALLE_AGRUPADORES a, TB_sublineas b, tb_lineas c where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_sli_sublinea_id = b.vcha_sli_sublinea_id and a.vcha_lin_linea_id = c.vcha_lin_linea_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(2).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
               rs.MoveNext:
            Wend
            rs.Close
            cmb_detalle_agrupadores(0).Enabled = False
            cmb_detalle_agrupadores(1).Enabled = True
            cmb_detalle_agrupadores(2).Enabled = True
            cmb_detalle_agrupadores(3).Enabled = False
            cmb_detalle_agrupadores(4).Enabled = False
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = True
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = True
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 4 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = True
            lv_detalle_agrupadores(4).Visible = False
            rs.Open "select distinct a.vcha_pro_producto_id,b.vcha_pro_nombre from TB_DETALLE_AGRUPADORES a, TB_productos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_pro_producto_id = b.vcha_pro_producto_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(3).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            cmb_detalle_agrupadores(0).Enabled = False
            cmb_detalle_agrupadores(1).Enabled = False
            cmb_detalle_agrupadores(2).Enabled = False
            cmb_detalle_agrupadores(3).Enabled = True
            cmb_detalle_agrupadores(4).Enabled = False
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = True
            lv_detalle_agrupadores(4).Visible = False
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = True
            opt_tipoagrupador(4).Value = False
            pro_textos
         End If
         If vartipoagrupador = 5 Then
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = True
            rs.Open "select distinct a.vcha_tar_tipo_articulo_id,b.vcha_tar_nombre from TB_DETALLE_AGRUPADORES a, TB_tipoarticulos b where a.vcha_agr_agrupador_id = '" & varagrupador & "' and a.vcha_tar_tipo_articulo_id = b.vcha_tar_tipo_articulo_id", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_detalle_agrupadores(4).ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext:
            Wend
            rs.Close
            cmb_detalle_agrupadores(0).Enabled = False
            cmb_detalle_agrupadores(1).Enabled = False
            cmb_detalle_agrupadores(2).Enabled = False
            cmb_detalle_agrupadores(3).Enabled = True
            cmb_detalle_agrupadores(4).Enabled = True
            lv_detalle_agrupadores(0).Visible = False
            lv_detalle_agrupadores(1).Visible = False
            lv_detalle_agrupadores(2).Visible = False
            lv_detalle_agrupadores(3).Visible = False
            lv_detalle_agrupadores(4).Visible = True
            opt_tipoagrupador(0).Value = False
            opt_tipoagrupador(1).Value = False
            opt_tipoagrupador(2).Value = False
            opt_tipoagrupador(3).Value = False
            opt_tipoagrupador(4).Value = True
            pro_textos
         End If
         rsaux2.MoveNext:
         Wend
      Else
         vartipoagrupador = 1
      End If
      rsaux2.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
   If opt_tipoagrupador(0).Value = True Then
      txt_detalle_agrupadores(0) = lv_detalle_agrupadores(0).selectedItem
      cmb_detalle_agrupadores(0) = Obtener_llave(cnn, rs, "TB_articulos", "VCHA_art_articulo_ID", txt_detalle_agrupadores(0), 1, "T")
      opt_tipoagrupador(0).Value = True
      opt_tipoagrupador(1).Value = False
      opt_tipoagrupador(2).Value = False
      opt_tipoagrupador(3).Value = False
      opt_tipoagrupador(4).Value = False
   End If
   If opt_tipoagrupador(1).Value = True Then
      txt_detalle_agrupadores(1) = lv_detalle_agrupadores(1).selectedItem
      cmb_detalle_agrupadores(1) = Obtener_llave(cnn, rs, "TB_lineas", "VCHA_lin_linea_ID", txt_detalle_agrupadores(1), 1, "T")
      opt_tipoagrupador(0).Value = False
      opt_tipoagrupador(1).Value = True
      opt_tipoagrupador(2).Value = False
      opt_tipoagrupador(3).Value = False
      opt_tipoagrupador(4).Value = False
   End If
   If opt_tipoagrupador(2).Value = True Then
      txt_detalle_agrupadores(1) = lv_detalle_agrupadores(2).selectedItem
      cmb_detalle_agrupadores(1) = Obtener_llave(cnn, rs, "TB_lineas", "VCHA_lin_linea_ID", txt_detalle_agrupadores(1), 1, "T")
      txt_detalle_agrupadores(2) = lv_detalle_agrupadores(2).selectedItem.SubItems(2)
      cmb_detalle_agrupadores(2) = Obtener_llave(cnn, rs, "TB_sublineas", "VCHA_sli_sublinea_ID", txt_detalle_agrupadores(2), 2, "T")
      opt_tipoagrupador(0).Value = False
      opt_tipoagrupador(1).Value = False
      opt_tipoagrupador(2).Value = True
      opt_tipoagrupador(3).Value = False
      opt_tipoagrupador(4).Value = False
   End If
   If opt_tipoagrupador(3).Value = True Then
      txt_detalle_agrupadores(3) = lv_detalle_agrupadores(3).selectedItem
      cmb_detalle_agrupadores(3) = Obtener_llave(cnn, rs, "TB_productos", "VCHA_pro_producto_ID", txt_detalle_agrupadores(3), 1, "T")
      opt_tipoagrupador(0).Value = False
      opt_tipoagrupador(1).Value = False
      opt_tipoagrupador(2).Value = False
      opt_tipoagrupador(3).Value = True
      opt_tipoagrupador(4).Value = False
   End If
   If opt_tipoagrupador(4).Value = True Then
      txt_detalle_agrupadores(3) = lv_detalle_agrupadores(3).selectedItem
      cmb_detalle_agrupadores(3) = Obtener_llave(cnn, rs, "TB_productos", "VCHA_pro_producto_ID", txt_detalle_agrupadores(3), 1, "T")
      txt_detalle_agrupadores(4) = lv_detalle_agrupadores(4).selectedItem
      cmb_detalle_agrupadores(4) = Obtener_llave(cnn, rs, "TB_tipoarticulos", "VCHA_tar_tipo_articulo_ID", txt_detalle_agrupadores(4), 1, "T")
      opt_tipoagrupador(0).Value = False
      opt_tipoagrupador(1).Value = False
      opt_tipoagrupador(2).Value = False
      opt_tipoagrupador(3).Value = False
      opt_tipoagrupador(4).Value = True
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
   If vartipoagrupador = 1 Then
      If var_modifica_registro_agrupador = False Then
         Set list_item = lv_detalle_agrupadores(0).ListItems.Add(, , txt_detalle_agrupadores(0))
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index) = txt_detalle_agrupadores(0)
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(0).ListItems.Item(lv_detalle_agrupadores(0).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 2 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(1).ListItems.Add(, , txt_detalle_agrupadores(1))
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index) = txt_detalle_agrupadores(1)
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(1).ListItems.Item(lv_detalle_agrupadores(1).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 3 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(2).ListItems.Add(, , txt_detalle_agrupadores(2))
         list_item.SubItems(1) = vardetallearticulo
         list_item.SubItems(2) = vardetallearticulo
         list_item.SubItems(3) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index) = txt_detalle_agrupadores(2)
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).ListSubItems(2) = vardetallearticulo
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).ListSubItems(3) = vardetallearticulo
         lv_detalle_agrupadores(2).ListItems.Item(lv_detalle_agrupadores(2).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 4 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(3).ListItems.Add(, , txt_detalle_agrupadores(3))
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index) = txt_detalle_agrupadores(3)
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(3).ListItems.Item(lv_detalle_agrupadores(3).selectedItem.Index).Selected = True
      End If
   End If
   If vartipoagrupador = 5 Then
      If var_modifica_registro = False Then
         Set list_item = lv_detalle_agrupadores(4).ListItems.Add(, , txt_detalle_agrupadores(4))
         list_item.SubItems(1) = vardetallearticulo
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index).Checked = False
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index) = txt_detalle_agrupadores(4)
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index).ListSubItems(1) = vardetallearticulo
         lv_detalle_agrupadores(4).ListItems.Item(lv_detalle_agrupadores(4).selectedItem.Index).Selected = True
      End If
   End If
End Sub

Private Sub txt_detalle_agrupadores_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_detalle_agrupadores_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub




