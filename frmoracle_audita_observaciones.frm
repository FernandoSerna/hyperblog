VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_audita_observaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observaciones de la auditoria"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   3480
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   9045
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2925
         Left            =   60
         TabIndex        =   5
         Top             =   465
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5159
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad Lector"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad Aduana"
            Object.Width           =   2593
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Observaciones"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   30
         TabIndex        =   4
         Top             =   135
         Width           =   8955
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   75
      TabIndex        =   0
      Top             =   3555
      Width           =   9045
      Begin VB.TextBox txt_observaciones 
         Height          =   2385
         Left            =   60
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   8910
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Observaciones"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   8955
      End
   End
End
Attribute VB_Name = "frmoracle_audita_observaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_observaciones_auditoria = ""
   rsaux12.Open "SELECT embarque, caja, codigo, description, cantidad_original, cantidad_auditada FROM XXVIA_TB_cAJAS_AUDITADAS a, xxvia_system_items_b where  EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " and codigo = segment1 and organization_id = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rsaux12.EOF
         Set list_item = Me.lv_lista.ListItems.Add(, , rsaux12!CODIGO)
         list_item.SubItems(1) = IIf(IsNull(rsaux12!Description), "", rsaux12!Description)
         list_item.SubItems(2) = IIf(IsNull(rsaux12!cantidad_original), "", rsaux12!cantidad_original)
         list_item.SubItems(3) = IIf(IsNull(rsaux12!cantidad_auditada), "", rsaux12!cantidad_auditada)
         rsaux12.MoveNext
   Wend
   rsaux12.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_observaciones_auditoria = Me.txt_observaciones
End Sub

