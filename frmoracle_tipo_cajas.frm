VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_tipo_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecion de caja"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_cajas 
      Height          =   3210
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   4605
      Begin MSComctlLib.ListView lv_lista 
         Height          =   3030
         Left            =   45
         TabIndex        =   1
         Top             =   135
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   5345
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
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   7408
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_tipo_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If var_cliente_costales <> "" Then
      Me.Caption = Me.Caption + " Costales"
   End If
   If var_cn_frontera <> "" Then
      Me.Caption = Me.Caption + " CN Frontera"
   End If
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If var_cliente_costales = "" Or var_cn_frontera <> "" Then
      If var_cn_frontera = "" Then
         rsaux10.Open "select * from tb_oracle_empaques order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
      Else
         rsaux11.Open "select * from TB_ORACLE_CN_FRONTERA where clave = '" + var_cn_frontera + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux11.EOF Then
            rsaux10.Open "select * from tb_oracle_empaques where exportaciones = 1 order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
         Else
            If var_unidad_organizacional = 90 Then
               rsaux10.Open "select * from tb_oracle_empaques where empaque <> 'CAJA BIASI' order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux10.Open "select * from tb_oracle_empaques order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
         rsaux11.Close
      End If
   Else
      rsaux11.Open "select * FROM TB_ORACLE_CLIENTES_COSTALES WHERE CLAVE = '" + var_cliente_costales + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux11.EOF Then
         rsaux10.Open "select * from tb_oracle_empaques where empaque like 'COSTAL%' OR  EMPAQUE = 'CAJA CORTINERO' OR EMPAQUE = 'CAJA BIASI' order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
      Else
         If var_unidad_organizacional = 90 Then
            rsaux10.Open "select * from tb_oracle_empaques where empaque <> 'CAJA BIASI' order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
         Else
            rsaux10.Open "select * from tb_oracle_empaques order by ORDEN, EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
      rsaux11.Close
      
   End If
   While Not rsaux10.EOF
         Set list_item = Me.lv_lista.ListItems.Add(, , IIf(IsNull(rsaux10(0).Value), "", rsaux10(0).Value))
         rsaux10.MoveNext
   Wend
   rsaux10.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.lv_lista.ListItems.Count > 0 Then
      var_nombre_caja = Me.lv_lista.selectedItem
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_nombre_caja = Me.lv_lista.selectedItem
      Unload Me
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
