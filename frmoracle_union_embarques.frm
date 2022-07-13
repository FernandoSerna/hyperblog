VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_union_embarques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Union de embarques"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      Picture         =   "frmoracle_union_embarques.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Movimiento"
      Top             =   120
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Picture         =   "frmoracle_union_embarques.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   120
      Width           =   330
   End
   Begin VB.TextBox txt_embarque 
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt_grupo 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.ListView lv_grupos 
      Height          =   4725
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8334
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
         Text            =   "Grupo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques 
      Height          =   4725
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8334
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
         Text            =   "Grupo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_union_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_buscar_Click()
   var_fecha = Date
   frmcalendario.mes = Date
   frmcalendario.Show 1
   var_fecha = var_fecha_general
   If IsDate(var_fecha) Then
      var_dia_s = CStr(Day(var_fecha))
      var_mes_s = CStr(Month(var_fecha))
      var_año_s = CStr(Year(var_fecha))
      If Len(var_dia_s) = 1 Then
         var_dia_s = "0" + var_dia_s
      End If
      If Len(var_mes_s) = 1 Then
         var_mes_s = "0" + var_mes_s
      End If
      If Len(var_año_s) = 2 Then
         var_año_s = "20" + var_año_s
      End If
      var_fecha_s1 = var_dia_s + "/" + var_mes_s + "/" + var_año_s
      var_fecha_s = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
      rs.Open "select distinct grupo, fecha_grupo from TB_ORACLE_GRUPOS_EMBARQUES where FECHA_GRUPO = " + var_fecha_s, cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Me.lv_embarques.ListItems.Clear
            Set list_item = lv_grupos.ListItems.Add(, , rs!grupo)
            list_item.SubItems(1) = var_fecha_s1
            rs.MoveNext
      Wend
      rs.Close
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_nuevo_Click()
    rs.Open "select isnull(max(cast(grupo as float)),0) + 1 as grupo from TB_ORACLE_GRUPOS_EMBARQUES", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       Me.txt_grupo = rs!grupo
    End If
    rs.Close

End Sub

Private Sub Form_Load()
      var_dia = CStr(Day(CDate(Now)))
      var_mes = CStr(Month(CDate(Now)))
      var_año = CStr(Year(CDate(Now)))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
     
     
     
     
      rs.Open "SELECT grupo, max(fecha_grupo) fecha_grupo FROM TB_ORACLE_GRUPOS_EMBARQUES WHERE fecha_grupo >= " + var_fecha + " and fecha_grupo < " + var_fecha + "+ 1 group by grupo", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
      
         While Not rs.EOF
               Set list_item = lv_grupos.ListItems.Add(, , rs!grupo)
               list_item.SubItems(1) = IIf(IsNull(rs!Fecha_grupo), "", rs!Fecha_grupo)
               rs.MoveNext
         Wend
         
      End If
      rs.Close

End Sub

Private Sub lv_embarques_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      var_si = MsgBox("¿Desea sacar el embarque del grupo?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "DELETE FROM TB_ORACLE_GRUPOS_EMBARQUES WHERE EMBARQUE = " + Me.lv_embarques.selectedItem, cnn, adOpenDynamic, adLockOptimistic
         lv_embarques.ListItems.Remove (lv_embarques.selectedItem.Index)
      End If
   End If
End Sub

Private Sub lv_grupos_GotFocus()
   If Me.lv_grupos.ListItems.Count > 0 Then
      Me.lv_embarques.ListItems.Clear
      Me.txt_grupo = Me.lv_grupos.selectedItem
      rs.Open "SELECT * FROM TB_ORACLE_GRUPOS_EMBARQUES WHERE GRUPO = '" + Me.lv_grupos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         
         While Not rs.EOF
               Set list_item = lv_embarques.ListItems.Add(, , rs!Embarque)
               list_item.SubItems(1) = IIf(IsNull(rs!Fecha_embarque), "", rs!Fecha_embarque)
               rs.MoveNext
         Wend
         
      End If
      rs.Close
   End If
End Sub

Private Sub lv_grupos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_grupos.ListItems.Count > 0 Then
      Me.lv_embarques.ListItems.Clear
      Me.txt_grupo = Me.lv_grupos.selectedItem
      rs.Open "SELECT * FROM TB_ORACLE_GRUPOS_EMBARQUES WHERE GRUPO = '" + Me.lv_grupos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
      
         While Not rs.EOF
               Set list_item = lv_embarques.ListItems.Add(, , rs!Embarque)
               list_item.SubItems(1) = IIf(IsNull(rs!Fecha_embarque), "", rs!Fecha_embarque)
               rs.MoveNext
         Wend
         
      End If
      rs.Close
   End If
   
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_grupo <> "" Then
         If IsNumeric(Me.txt_embarque) Then
            rs.Open "SELECT * FROM TB_ORACLE_GRUPOS_EMBARQUES WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rsaux.Open "SELECT * FROM XXVIA_tB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_dia = CStr(Day(CDate(rsaux!FECHA_INICIO)))
                  var_mes = CStr(Month(CDate(rsaux!FECHA_INICIO)))
                  var_año = CStr(Year(CDate(rsaux!FECHA_INICIO)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  var_fecha_s = var_dia + "/" + var_mes + "/" + var_año
                  
                  var_dia = CStr(Day(CDate(Now)))
                  var_mes = CStr(Month(CDate(Now)))
                  var_año = CStr(Year(CDate(Now)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_g = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
                  
                  rsaux1.Open "INSERT INTO TB_ORACLE_GRUPOS_EMBARQUES (GRUPO, FECHA_GRUPO, EMBARQUE, FECHA_EMBARQUE) VALUES ('" + Me.txt_grupo + "'," + var_fecha_g + ",'" + Me.txt_embarque + "'," + var_fecha + ")", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = lv_embarques.ListItems.Add(, , Me.txt_embarque)
                  list_item.SubItems(1) = var_fecha_s
               Else
                  MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            Else
               MsgBox "El embarque ya se encuentra en el grupo " + rs!grupo, vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_grupo_Change()
   Me.lv_embarques.ListItems.Clear
End Sub

Private Sub txt_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
   var_fecha = Date
   If KeyCode = 116 Then
      frmcalendario.mes = Date
      frmcalendario.Show 1
      var_fecha = var_fecha_general
   End If
   If IsDate(var_fecha) Then
      var_dia_s = CStr(Day(var_fecha))
      var_mes_s = CStr(Month(var_fecha))
      var_año_s = CStr(Year(var_fecha))
      If Len(var_dia_s) = 1 Then
         var_dia_s = "0" + var_dia_s
      End If
      If Len(var_mes_s) = 1 Then
         var_mes_s = "0" + var_mes_s
      End If
      If Len(var_año_s) = 2 Then
         var_año_s = "20" + var_dia_s
      End If
      var_fecha_s1 = var_dia_s + "/" + var_mes_s + "/" + var_s
      var_fecha_s = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
      rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where FECHA_GRUPO = " + var_fecha_s, cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_grupos.ListItems.Add(, , rs!grupo)
            list_item.SubItems(1) = var_fecha_s1
            rs.MoveNext
      Wend
      rs.Close
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_grupo_KeyPress(KeyAscii As Integer)
   If IsNumeric(Me.txt_grupo) Then
      rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where grupo = '" + Me.txt_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.lv_grupos.ListItems.Clear
         Me.lv_embarques.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_embarques.ListItems.Add(, , rs!Embarque)
               list_item.SubItems(1) = rs!Fecha_embarque
               rs.MoveNext
         Wend
      Else
         MsgBox "El grupo no existe", vbOKOnly, "ATENCION"
         Me.txt_grupo.Text = ""
      End If
      rs.Close
   End If
End Sub
