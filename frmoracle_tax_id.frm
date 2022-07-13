VERSION 5.00
Begin VB.Form frmoracle_tax_id 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TAX ID"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   345
      Picture         =   "frmoracle_tax_id.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   30
      Picture         =   "frmoracle_tax_id.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7875
      Picture         =   "frmoracle_tax_id.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   10
      Top             =   300
      Width           =   8145
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   45
      TabIndex        =   7
      Top             =   390
      Width           =   8130
      Begin VB.TextBox txt_observaciones 
         Height          =   1095
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   6585
      End
      Begin VB.TextBox txt_tax_id 
         Height          =   420
         Left            =   1320
         TabIndex        =   5
         Top             =   630
         Width           =   2430
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   405
         Left            =   2700
         TabIndex        =   4
         Top             =   180
         Width           =   5190
      End
      Begin VB.TextBox txt_cliente 
         Height          =   405
         Left            =   1305
         TabIndex        =   3
         Top             =   180
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1185
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TAX ID:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmoracle_tax_id"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_nuevo_Click()
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_tax_id = ""
   Me.txt_observaciones = ""
   Me.txt_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   If Me.txt_nombre_cliente <> "" Then
      strconsulta = "SELECT * FROM XXVIA_TB_TAX_ID WHERE CLIENTE = ? "
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_cliente)
           .Parameters.Append parametro
      End With
      Set rs = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rs.EOF Then
         strconsulta = "UPDATE XXVIA_TB_TAX_ID SET TAX_ID = ?, observaciones = ? WHERE CLIENTE = ? "
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_tax_id)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.txt_observaciones)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_cliente)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
      Else
         strconsulta = "INSERT INTO  XXVIA_TB_TAX_ID  (CLIENTE, TAX_ID, observaciones) VALUES  (?, ?, ?)"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_cliente)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_tax_id)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 2000, Me.txt_observaciones)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "No se a indicado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
    Top = 3000
    Left = 1500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre_cliente.SetFocus
   End If
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(Me.txt_cliente) <> "" Then
      strconsulta = "SELECT PARTY_SITE_NUMBER AS CLIENTE, RAZON_SOCIAL_CLIENTE  FROM XXVIA_VW_CLIENTES_BCP WHERE PARTY_SITE_NUMBER = ? AND SITE_USE_CODE = 'BILL_TO'"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_cliente)
           .Parameters.Append parametro
      End With
      Set rs = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         strconsulta = "SELECT * FROM XXVIA_TB_TAX_ID WHERE CLIENTE = ? "
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_cliente)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux.EOF Then
            Me.txt_tax_id = IIf(IsNull(rsaux!TAX_ID), "", rsaux!TAX_ID)
            Me.txt_observaciones = IIf(IsNull(rsaux!observaciones), "", rsaux!observaciones)
         Else
            Me.txt_tax_id = ""
            Me.txt_observaciones = ""
         End If
         rsaux.Close
      Else
         MsgBox "El cliente no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_cliente = ""
      Me.txt_tax_id = ""
      Me.txt_observaciones = ""
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tax_id.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_observaciones_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.com_guardar.SetFocus
   End If
End Sub

Private Sub txt_tax_id_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_observaciones.SetFocus
   End If
End Sub
