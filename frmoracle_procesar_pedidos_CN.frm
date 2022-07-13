VERSION 5.00
Begin VB.Form frmoracle_procesar_pedidos_CN 
   Caption         =   "Procesar notas con error al cargar"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   3675
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_pedido 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmoracle_procesar_pedidos_CN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         If cnnicg_sql.State = 1 Then
            cnnicg_sql.Close
         End If
         rsaux.Open "select distinct vcha_nota_envio, numb_status from xxpos.xxvia_tb_icg_tran_cedis_tienda where vcha_nota_envio = '" + Me.txt_pedido + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            If IIf(IsNull(rsaux!numb_status), 0, rsaux!numb_status) = 0 Then
               cnnicg_sql.Open "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
               rsaux9.Open "exec vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_pedido) + "'", cnnicg_sql, adOpenDynamic, adLockOptimistic
               rsaux9.Open "call xxpos.xxvia_pk_motor_logistico.xxvia_sp_senales_eviandas_a_cn (" + Me.txt_pedido + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux9.Open "UPDATE TB_ORACLE_ERRORES_NOTAS_CN SET PROCESADA=1 WHERE PEDIDO = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "La nota ya habia sido cargada", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La nota no a sido cargada", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
   End If
End Sub
