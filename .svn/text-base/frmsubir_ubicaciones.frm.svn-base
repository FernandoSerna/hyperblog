VERSION 5.00
Begin VB.Form frmsubir_ubicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subir ubicaciones"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Subir ubicaciones"
      Height          =   795
      Left            =   165
      TabIndex        =   0
      Top             =   285
      Width           =   4365
   End
End
Attribute VB_Name = "frmsubir_ubicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   txt_ruta = "c:\ubicaciones.xls"
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & txt_ruta
   rsaux2.Open "SELECT * FROM [UBICACIONES$]", strConnectionString
   While Not rsaux2.EOF
         If Not IsNull(rsaux2!codigo) Then
            If Not IsNull(rsaux2!UBICACION) Then
               rsaux3.Open "select * from tb_ubicaciones_almacen where vcha_Alm_almacen_id = '8' and vcha_art_articulo_id = '" + Trim(rsaux2!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  rsaux.Open "update tb_ubicaciones_almacen set vcha_ubi_ubicacion_2 = '" + Trim(rsaux2!UBICACION) + "' where vcha_Art_articulo_id = '" + Trim(rsaux2!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux.Open "insert into tb_ubicaciones_almacen (vcha_alm_almacen_id, vcha_Art_articulo_id, vcha_ubi_ubicacion_2) values ('8','" + Trim(rsaux2!codigo) + "','" + Trim(rsaux2!UBICACION) + "')", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux3.Close
            End If
         End If
         rsaux2.MoveNext
    Wend
    rsaux2.Close
    MsgBox "Se a terminado de cargar las ubicaciones "
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub
