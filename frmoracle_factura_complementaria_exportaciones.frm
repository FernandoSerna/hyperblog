VERSION 5.00
Begin VB.Form frmoracle_factura_complementaria_exportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factura complementaria"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_factura 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   337
      Width           =   825
   End
End
Attribute VB_Name = "frmoracle_factura_complementaria_exportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim VAR_TIPO_LISTA As Integer

Private Sub Form_Load()
   Top = 3000
   Left = 4100
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_factura) Then
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.Open "SELECT * FROM RA_CUSTOMER_TRX_ALL WHERE TRX_NUMBER = '" + Me.txt_factura + "' AND CUST_TRX_TYPE_ID IN (1244, 1039)", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "SELECT * FROM XXVIA_TB_COMP_FAC WHERE VCHA_NUMERO_FACTURA = '" + Me.txt_factura + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               MsgBox "La factura complementaria ya existe. Factura: " + rsaux!vcha_referencia, vbOKOnly, "ATENCION"
            Else
               rsaux1.Open " call xxvia_sp_complemento_factura ('" + Me.txt_factura + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux1.Open "SELECT * FROM XXVIA_TB_COMP_FAC WHERE VCHA_NUMERO_FACTURA = '" + Me.txt_factura + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rsaux1.EOF Then
                  MsgBox "La factura no contiene artículos complementarios", vbOKOnly, "ATENCION"
               Else
                  MsgBox "Se a generado la factura " + rsaux1!vcha_referencia + " favor de correr el concurrente Eflow", vbOKOnly, "ATENCION"
               End If
               rsaux1.Close
            End If
            rsaux.Close
         End If
         rs.Close
         
      Else
         MsgBox "Número de factura incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
