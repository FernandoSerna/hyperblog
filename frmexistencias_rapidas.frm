VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmexistencias_rapidas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   3210
      Left            =   60
      TabIndex        =   1
      Top             =   1350
      Width           =   8700
      Begin MSComctlLib.ListView lv_lista 
         Height          =   3000
         Left            =   75
         TabIndex        =   5
         Top             =   135
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   5292
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Almacén"
            Object.Width           =   11377
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   8700
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1095
         TabIndex        =   4
         Top             =   765
         Width           =   7515
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1095
         TabIndex        =   3
         Top             =   255
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción:"
         Height          =   225
         Left            =   165
         TabIndex        =   6
         Top             =   900
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   225
         Left            =   270
         TabIndex        =   2
         Top             =   383
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmexistencias_rapidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.lv_lista.ListItems.Clear
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      If Len(Me.txt_codigo) <= 5 Then
         For var_j = Len(Me.txt_codigo) To 7
            Me.txt_codigo = "0" + Me.txt_codigo
         Next var_j
      End If
      rs.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!Description), "", rs!Description)
         strconsulta = "SELECT b.description, cantmano FROM Xxvia_vw_existencias_inv a, mtl_secondary_inventories b WHERE a.ORGANIZATION_ID = ? AND SEGMENT1 = ? and subinventory_code not like '%TD%' and a.organization_id = b.organization_id and a.subinventory_code = b.secondary_inventory_name"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing


         'rsaux.Open "SELECT b.description, cantmano FROM Xxvia_vw_existencias_inv a, mtl_secondary_inventories b WHERE a.ORGANIZATION_ID = " + var_unidad_organizacional + " AND SEGMENT1 = '" + Me.txt_codigo + "' and subinventory_code not like '%TD%' and a.organization_id = b.organization_id and a.subinventory_code = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  Set list_item = Me.lv_lista.ListItems.Add(, , rsaux!Description)
                  list_item.SubItems(1) = Format(rsaux!CANTMANO)
                  rsaux.MoveNext
            Wend
         
         Else
            Me.lv_lista.ListItems.Clear
         End If
         rsaux.Close
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_descripcion = ""
      Me.lv_lista.ListItems.Clear
   End If
End Sub
