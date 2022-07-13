VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_agrupacion_bultos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agrupación de bultos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_pedido 
      Height          =   285
      Left            =   3960
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt_embarque 
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6600
      Picture         =   "frmoracle_agrupacion_bultos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   120
      Picture         =   "frmoracle_agrupacion_bultos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6855
      Begin VB.TextBox txt_caja_unir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   2895
      End
      Begin MSComctlLib.ListView lv_cajas 
         Height          =   2805
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   4948
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   11289
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Caja a unir:"
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
         TabIndex        =   7
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Cajas a unir"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   6810
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6855
   End
   Begin VB.TextBox txt_caja_padre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Caja Padre:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1635
   End
End
Attribute VB_Name = "frmoracle_agrupacion_bultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   var_si = MsgBox("¿Desea agrupar los bultos?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar agrupar los bultos", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To Me.lv_cajas.ListItems.Count
             Me.lv_cajas.ListItems.Item(var_j).Selected = True
             rs.Open "update TB_ORACLE_CAJAS_ADUANA set caja_anterior = '" + Me.lv_cajas.selectedItem + "', pedido_anterior = '" + Me.txt_pedido + "', embarque_anterior = '" + Me.txt_embarque + "'  WHERE EMBARQUE = " + Me.txt_embarque + " and pedido = " + Me.txt_pedido + " and caja = '" + Me.lv_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
             rs.Open "update TB_ORACLE_CAJAS_ADUANA set embarque = 0, PEDIDO  = 0 WHERE EMBARQUE = " + Me.txt_embarque + " and pedido = " + Me.txt_pedido + " and caja = '" + Me.lv_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
             var_caja_actual = Mid(Me.txt_caja_padre, 8, 3)
             var_caja = CDbl(Mid(Me.lv_cajas.selectedItem, 8, 3))
             strconsulta = "update xxvia_tb_salidas_cajas set inte_paq_caja_Anterior = ?, inte_paq_caja = ? where inte_emb_embarque = ? and source_header_number = ? and inte_paq_caja = ?"
             With comandoORA
                  .ActiveConnection = cnnoracle_4
                  .CommandType = adCmdText
                  .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja))
                  .Parameters.Append parametro
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja_actual))
                  .Parameters.Append parametro
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(Me.txt_embarque))
                  .Parameters.Append parametro
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(Me.txt_pedido))
                  .Parameters.Append parametro
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja))
                  .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing

         Next var_j
      End If
   End If
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_caja_padre_Change()
   Me.txt_caja_unir = ""
   Me.txt_embarque = ""
   Me.txt_pedido = ""
   Me.lv_cajas.ListItems.Clear
End Sub

Private Sub txt_caja_padre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      rs.Open "select * from TB_ORACLE_CAJAS_ADUANA where CAJA =  '" + Me.txt_caja_padre + "' AND EMBARQUE = SUBSTRING('" + Me.txt_caja_padre + "',2,6)", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_caja_unir.SetFocus
         Me.txt_embarque = rs!Embarque
         Me.txt_pedido = rs!pedido
      Else
         MsgBox "El bulto no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_caja_unir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rsaux.Open "select * from TB_ORACLE_CAJAS_ADUANA where CAJA_ANTERIOR =  '" + Me.txt_caja_unir + "' AND EMBARQUE = SUBSTRING('" + Me.txt_caja_unir + "',2,6)", cnn, adOpenDynamic, adLockOptimistic
      If rsaux.EOF Then
         var_posible = 1
         For var_j = 1 To Me.lv_cajas.ListItems.Count
             Me.lv_cajas.ListItems.Item(var_j).Selected = True
             If Me.lv_cajas.selectedItem = Me.txt_caja_unir Then
                var_posible = 0
             End If
         Next var_j
         If var_posible = 1 Then
            rs.Open "select * from TB_ORACLE_CAJAS_ADUANA where CAJA =  '" + Me.txt_caja_unir + "' AND EMBARQUE = SUBSTRING('" + Me.txt_caja_unir + "',2,6)", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_embarque = rs!Embarque
               var_pedido = rs!pedido
               If var_pedido = Me.txt_pedido Then
                  Set list_item = Me.lv_cajas.ListItems.Add(, , Me.txt_caja_unir)
                  Me.txt_caja_unir = ""
               Else
                  MsgBox "La bulto no corresponde al pedido del bulto padre.", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El bulto no existe o no pertenece al embarque del bulto padre.", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
         End If
      Else
         MsgBox "El bulto ya fue agrupado al bulto " + rsaux!caja_actual, vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   End If

End Sub

