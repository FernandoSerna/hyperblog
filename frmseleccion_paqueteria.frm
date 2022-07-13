VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmseleccion_paqueteria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de paqueteria"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_lista 
      Height          =   1995
      Left            =   300
      TabIndex        =   12
      Top             =   15
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1455
         Left            =   30
         TabIndex        =   13
         Top             =   495
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2566
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmseleccion_paqueteria.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmseleccion_paqueteria.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   45
      TabIndex        =   11
      Top             =   360
      Width           =   6090
   End
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   60
      TabIndex        =   7
      Top             =   375
      Width           =   6015
      Begin VB.TextBox txt_guia 
         Height          =   360
         Left            =   1050
         TabIndex        =   4
         Top             =   1020
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txt_nombre_caja 
         Height          =   360
         Left            =   1725
         TabIndex        =   3
         Top             =   630
         Width           =   4185
      End
      Begin VB.TextBox txt_caja 
         Height          =   360
         Left            =   1050
         TabIndex        =   2
         Top             =   630
         Width           =   660
      End
      Begin VB.TextBox txt_nombre_paqueteria 
         Height          =   360
         Left            =   1725
         TabIndex        =   1
         Top             =   255
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.TextBox txt_paqueteria 
         Height          =   360
         Left            =   1050
         TabIndex        =   0
         Top             =   255
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1110
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   713
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paqueteria:"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   338
         Visible         =   0   'False
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmseleccion_paqueteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_TIPO_LISTA As Integer
Private Sub cmd_aceptar_Click()
   If Me.txt_guia <> "" Then
      If Me.txt_paqueteria <> "" Then
         If Me.txt_caja <> "" Then
            var_paqueteria = Me.txt_paqueteria
            var_nombre_paqueteria = Me.txt_nombre_paqueteria
            var_tamaño_caja = Me.txt_caja
            var_nombre_caja = Me.txt_nombre_caja
            var_guia = Me.txt_guia
            Unload Me
         Else
            var_si_asignacion_paqueteria = 0
            var_tamaño_caja = ""
            var_nombre_caja = ""
            var_paqueteria = Me.txt_paqueteria
            var_nombre_paqueteria = Me.txt_nombre_paqueteria
            var_tamaño_caja = Me.txt_caja
            var_nombre_caja = Me.txt_nombre_caja
            var_guia = Me.txt_guia
            Unload Me
         End If
      Else
         var_si_asignacion_paqueteria = 0
         var_paqueteria = ""
         var_nombre_paqueteria = ""
         var_paqueteria = Me.txt_paqueteria
         var_nombre_paqueteria = Me.txt_nombre_paqueteria
         var_tamaño_caja = Me.txt_caja
         var_nombre_caja = Me.txt_nombre_caja
         var_guia = Me.txt_guia
         Unload Me
      End If
   Else
      var_si_asignacion_paqueteria = 0
      var_guia = ""
      var_paqueteria = Me.txt_paqueteria
      var_nombre_paqueteria = Me.txt_nombre_paqueteria
      var_tamaño_caja = Me.txt_caja
      var_nombre_caja = Me.txt_nombre_caja
      var_guia = Me.txt_guia
      Unload Me
   End If
End Sub

Private Sub cmd_cancelar_Click()
   var_nombre_paqueteria = ""
   var_guia = ""
   var_paqueteria = ""
   var_nombre_caja = ""
   var_caja = ""
   var_si_asignacion_paqueteria = 0
   Unload Me
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Me.txt_caja = var_tamaño_caja
   Me.txt_nombre_caja = var_nombre_caja
   Me.txt_paqueteria = var_paqueteria
   Me.txt_nombre_paqueteria = var_nombre_paqueteria
   Me.txt_guia = var_guia
   If Me.txt_paqueteria <> "" Then
      Me.txt_paqueteria.Enabled = False
      Me.txt_nombre_paqueteria.Enabled = False
   Else
      Me.txt_paqueteria.Enabled = True
      Me.txt_nombre_paqueteria.Enabled = True
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If VAR_TIPO_LISTA = 1 Then
         Me.txt_paqueteria = Me.lv_lista.selectedItem
         Me.txt_nombre_paqueteria = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_paqueteria.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         Me.txt_caja = Me.lv_lista.selectedItem
         Me.txt_nombre_caja = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_caja.SetFocus
      End If
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      If VAR_TIPO_LISTA = 1 Then
         Me.txt_paqueteria.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         Me.txt_caja.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_caja_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rsaux9.Open "select DISTINCT VCHA_cAJ_cAJA_ID, VCHA_cAJ_NOMBRE from VW_PRECIOS_PAQUETERIA_SID WHERE VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "' order by vcha_CAJ_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux9.EOF Then
         While Not rsaux9.EOF
               Set list_item = lv_lista.ListItems.Add(, , rsaux9!vcha_caj_caja_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux9!VCHA_CAJ_NOMBRE), "", rsaux9!VCHA_CAJ_NOMBRE)
               rsaux9.MoveNext
         Wend
         rsaux9.Close
         lbl_lista = "CAJAS"
         VAR_TIPO_LISTA = 2
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         rsaux9.Close
         MsgBox "No se a indicado la paquetería de la orden de surtido", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_caja_LostFocus()
   If Me.txt_caja <> "" Then
      rsaux9.Open "SELECT * FROM VW_PRECIOS_PAQUETERIA_sid WHERE VCHA_CAJ_CAJA_ID = '" + Me.txt_caja + "' AND VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux9.EOF Then
         Me.txt_nombre_caja = IIf(IsNull(rsaux9!VCHA_CAJ_NOMBRE), "", rsaux9!VCHA_CAJ_NOMBRE)
      Else
         MsgBox "Clave de caja incorrecta", vbOKOnly, "ATENCION"
         Me.txt_caja = ""
         Me.txt_nombre_caja = ""
      End If
      rsaux9.Close
   Else
      Me.txt_nombre_caja = ""
   End If
End Sub

Private Sub txt_guia_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_caja_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rsaux9.Open "select DISTINCT VCHA_cAJ_cAJA_ID, VCHA_cAJ_NOMBRE from VW_PRECIOS_PAQUETERIA_SID WHERE VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "' order by vcha_CAJ_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux9.EOF
            Set list_item = lv_lista.ListItems.Add(, , rsaux9!vcha_caj_caja_id)
            list_item.SubItems(1) = IIf(IsNull(rsaux9!VCHA_CAJ_NOMBRE), "", rsaux9!VCHA_CAJ_NOMBRE)
            rsaux9.MoveNext
      Wend
      rsaux9.Close
      lbl_lista = "CAJAS"
      VAR_TIPO_LISTA = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_paqueteria_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rsaux9.Open "select * from tb_paqueteria order by vcha_paq_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux9.EOF
            Set list_item = lv_lista.ListItems.Add(, , rsaux9!vcha_paq_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rsaux9!vcha_paq_nombre), "", rsaux9!vcha_paq_nombre)
            rsaux9.MoveNext
      Wend
      rsaux9.Close
      lbl_lista = "PAQUETERIAS"
      VAR_TIPO_LISTA = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_paqueteria_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_paqueteria_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      
      rsaux9.Open "select * from tb_paqueteria order by vcha_paq_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux9.EOF
            Set list_item = lv_lista.ListItems.Add(, , rsaux9!vcha_paq_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rsaux9!vcha_paq_nombre), "", rsaux9!vcha_paq_nombre)
            rsaux9.MoveNext
      Wend
      rsaux9.Close
      lbl_lista = "PAQUETERIAS"
      VAR_TIPO_LISTA = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_paqueteria_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_paqueteria_LostFocus()
   If Me.txt_paqueteria <> "" Then
      rsaux9.Open "select * from tb_paqueteria where vcha_paq_clave_id = '" + Me.txt_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux9.EOF Then
         Me.txt_nombre_paqueteria = IIf(IsNull(rsaux9!vcha_paq_nombre), "", rsaux9!vcha_paq_nombre)
      Else
         MsgBox "Clave de paqueteria incorrecta", vbOKOnly, "ATENCION"
         Me.txt_nombre_paqueteria = ""
         Me.txt_paqueteria = ""
      End If
      rsaux9.Close
   Else
      Me.txt_nombre_paqueteria = ""
   End If
End Sub

