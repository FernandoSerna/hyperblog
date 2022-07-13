VERSION 5.00
Begin VB.Form frmoracle_validador_codigos_barras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validador de códigos de barra"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1740
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   11460
      Begin VB.TextBox txt_oracle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3300
         TabIndex        =   6
         Top             =   1185
         Width           =   2415
      End
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3300
         TabIndex        =   4
         Top             =   735
         Width           =   8055
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3300
         TabIndex        =   2
         Top             =   150
         Width           =   4260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código de Oracle:"
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
         TabIndex        =   5
         Top             =   1245
         Width           =   2580
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
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
         TabIndex        =   3
         Top             =   795
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código de barras:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   3165
      End
   End
End
Attribute VB_Name = "frmoracle_validador_codigos_barras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter



Private Sub Form_Load()
   Top = 2500
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_oracle = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT, a.cantidad FROM (select INVENTORY_ITEM_ID, description, cross_reference, nvl(attribute1,1) as cantidad from c) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
           .Parameters.Append parametro
      End With
      Set rsaux8 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux8.EOF Then
         Me.txt_oracle = rsaux8!SEGMENT1
         Me.txt_descripcion = IIf(IsNull(rsaux8!Description), "", rsaux8!Description)
      Else
         Me.txt_oracle = ""
         Me.txt_descripcion = ""
         MsgBox "Código de barras incorrecto", vbOKOnly, "ATENCION"
      End If
      rsaux8.Close
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT, a.cantidad FROM (select INVENTORY_ITEM_ID, description, cross_reference, nvl(attribute1,1) as cantidad from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
        .Parameters.Append parametro
        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
        .Parameters.Append parametro
   End With
   Set rsaux8 = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   If Not rsaux8.EOF Then
      Me.txt_oracle = rsaux8!SEGMENT1
      Me.txt_descripcion = IIf(IsNull(rsaux8!Description), "", rsaux8!Description)
   Else
      Me.txt_oracle = ""
      Me.txt_descripcion = ""
      MsgBox "Código de barras incorrecto", vbOKOnly, "ATENCION"
   End If
   rsaux8.Close
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_oracle.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_oracle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
