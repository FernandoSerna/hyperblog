VERSION 5.00
Begin VB.Form form_WS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar webservice"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   7230
      TabIndex        =   5
      Top             =   1455
      Width           =   1320
   End
   Begin VB.CommandButton consultar 
      Caption         =   "consultar"
      Height          =   540
      Left            =   7290
      TabIndex        =   4
      Top             =   2430
      Width           =   1305
   End
   Begin VB.TextBox wsdl 
      Height          =   435
      Left            =   4185
      TabIndex        =   3
      Top             =   180
      Width           =   4440
   End
   Begin VB.TextBox action 
      Height          =   390
      Left            =   4200
      TabIndex        =   2
      Top             =   810
      Width           =   4380
   End
   Begin VB.TextBox soap 
      Height          =   2835
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   165
      Width           =   3915
   End
   Begin VB.TextBox resultado 
      Height          =   2235
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3240
      Width           =   8385
   End
End
Attribute VB_Name = "form_WS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private WithEvents clase1 As DllHolaMundo.Class1
Dim xmlResponse As MSXML2.DOMDocument30
Dim strSoap As String
Dim strSOAPAction As String
Dim strWsdl As String





Private Sub Command1_Click()
Dim clnt As New SoapClient30
Dim str As String
clnt.MSSoapInit "http://intranet/WSOracle/wsInterfaceOM.asmx?wsdl"
str = clnt.crear_embarque("11111111", "11111111")
MsgBox str
End Sub

Private Sub consultar_Click()

strSoap = soap.Text
strSOAPAction = action.Text
strWsdl = wsdl.Text
If invokewebservice(strSoap, strSOAPAction, strWsdl, xmlResponse) Then
   resultado.Text = xmlResponseXML
Else
   resultado.Text = "Error"
End If
Set xmlResponse = Nothing
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 2000
    'Set clase1 = New ClassLibrary3.Class1
    '
    'txtQueSaludo = "el Guille"
    'lblInfo = clase1.ToString
    '

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
