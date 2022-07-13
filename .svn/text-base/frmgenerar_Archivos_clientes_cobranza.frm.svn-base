VERSION 5.00
Begin VB.Form frmgenerar_Archivos_clientes_cobranza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar archivos de clientes de cobranza"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_generar_Archivos 
      Caption         =   "Generar archivos"
      Height          =   795
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   4365
   End
End
Attribute VB_Name = "frmgenerar_Archivos_clientes_cobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ruta As String
Dim var_tabla As ADODB.Connection
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub cmd_generar_Archivos_Click()
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   rs.Open "select * from tbclient", var_tabla, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub
