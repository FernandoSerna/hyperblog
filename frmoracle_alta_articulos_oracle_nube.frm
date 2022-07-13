VERSION 5.00
Begin VB.Form frmoracle_alta_articulos_oracle_nube 
   Caption         =   "Alta codigos Nube a SID"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_uom 
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txt_clasificacion_sat 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txt_equivalencia 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txt_descripcion 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7455
   End
   Begin VB.TextBox txt_codigo 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmoracle_alta_articulos_oracle_nube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Me.txt_codigo <> "" Then
          rs.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, b.UNIT_WEIGHT, nvl(a.attribute1,1) as cantidad, clasificacionsat, uom_sat FROM mtl_cross_references_b A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND segment1 = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             Me.txt_descripcion = rs!Description
             Me.txt_equivalencia = rs!cross_reference
             Me.txt_clasificacion_sat = rs!CLASIFICACIONSAT
             Me.txt_uom = RE!UOM_SAT
             
          Else
             MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
             Me.txt_descripcion = ""
             Me.txt_equivalencia = ""
          End If
          
          
       Else
          Me.txt_descripcion = ""
          Me.txt_equivalencia = ""
       End If
    End If
End Sub
