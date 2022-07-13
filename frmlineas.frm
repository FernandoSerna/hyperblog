VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlineas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lineas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmlineas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   285
      Left            =   4710
      TabIndex        =   30
      Top             =   45
      Width           =   585
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   195
      Left            =   4320
      TabIndex        =   29
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   195
      Left            =   3930
      TabIndex        =   28
      Top             =   30
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   3465
      TabIndex        =   27
      Top             =   30
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   3015
      TabIndex        =   26
      Top             =   30
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   240
      Left            =   2535
      TabIndex        =   24
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   2040
      TabIndex        =   23
      Top             =   45
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmlineas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      Picture         =   "frmlineas.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmlineas.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmlineas.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmlineas.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmlineas.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2610
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   13
      Top             =   1770
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1785
         TabIndex        =   17
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3735
         TabIndex        =   22
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de linea:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   195
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Lineas "
      Height          =   1350
      Left            =   150
      TabIndex        =   7
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_capacidad 
         Height          =   315
         Left            =   3855
         MaxLength       =   3
         TabIndex        =   10
         Top             =   945
         Width           =   615
      End
      Begin VB.TextBox txt_lineas 
         Height          =   315
         Index           =   2
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   9
         Top             =   915
         Width           =   615
      End
      Begin VB.TextBox txt_lineas 
         Height          =   315
         Index           =   1
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   570
         Width           =   4155
      End
      Begin VB.TextBox txt_lineas 
         Height          =   315
         Index           =   0
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   6
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Capacidad en m³:"
         Height          =   195
         Index           =   4
         Left            =   2445
         TabIndex        =   25
         Top             =   1005
         Width           =   1245
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   21
         Top             =   975
         Width           =   120
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Compreción:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   20
         Top             =   975
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   630
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   285
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4905
      Left            =   150
      TabIndex        =   15
      Top             =   2325
      Width           =   5655
      Begin MSComctlLib.ListView lv_lineas 
         Height          =   4710
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8308
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "compresion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Capacidad"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3075
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   16
      Top             =   285
      Visible         =   0   'False
      Width           =   5655
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlineas.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmlineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_lineas As Integer
Dim bitacora As Boolean



Private Sub cmd_deshacer_Click()
       Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      Call pro_elimina_lineas
      rs.Open "select * from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   If Me.txt_lineas(2).Text = "" Then
      Me.txt_lineas(2).Text = 0
   End If
   var_posible = True
   If var_modifica_registro_linea = False Then
      rs.Open "select * from tb_lineas where vcha_lin_linea_id = '" + Me.txt_lineas(0) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
      var_opcion_seguridad = 2
      var_acepta_seguridad = 1
      If var_global_permiso3 = 1 Then
         var_acepta_seguridad = 2
         If var_global_permiso4 = 1 Then
            frmpasswords2.Show 1
         Else
            frmpasswords.Show 1
         End If
      End If
      If var_acepta_seguridad = 1 Then
         Call pro_guardar_lineas
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
            cmd_guardar.Enabled = False
            cmd_deshacer.Enabled = False
            cmd_eliminar.Enabled = False
         Else
            cmd_guardar.Enabled = True
            cmd_deshacer.Enabled = True
            cmd_eliminar.Enabled = True
         End If
         rs.Close
      End If
   Else
      MsgBox "Clave de liena ya existe", vbOKOnly, "ATENCION"
   End If
   If var_clave_usuario_global = "U0000000182" Then
      Me.txt_lineas(0).Enabled = False
      Me.txt_lineas(1).Enabled = False
      Me.txt_lineas(2).Enabled = False
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_lineas, "LISTADO DE lineas")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   If var_clave_usuario_global = "U0000000182" Then
      Me.txt_lineas(0).Enabled = False
      Me.txt_lineas(1).Enabled = False
      Me.txt_lineas(2).Enabled = False
   Else
      Call pro_limpiatextos(Me)
      txt_lineas(0).Enabled = True
      txt_lineas(0).SetFocus: var_modifica_registro_linea = False
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
  End If
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_linea = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
End Sub





Private Sub Command1_Click()
    

Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN

Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
Const vbAPINull As Long = 0&                         ' Puntero NULL

' se elimina
Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
sDsnName = "DSN=sqlsistema"
sDriver = "SQL Server"
dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

'se crea
sDsnName = "sqlsistema"
sDescription = "sqlsistema"
sDriver = "SQL Server"
sAttributes = "DSN=" & sDsnName & Chr(0)
sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
strAttributes = strAttributes & "UID=sa" & Chr$(0)
strAttributes = strAttributes & "PWD=elia" & Chr$(0)
dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)







End Sub

Private Sub Command2_Click()
Dim dl As Long                        ' Valor devuelto por la función API
Dim sDriver As String              ' Nombre del controlador
Dim sDsnName As String         ' Nombre del DSN

Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema

' Establecemos los atributos necesarios

' CUIDADO: no dejar espacios en blanco entre el parámetro
' «DSN», el signo igual y el nombre del DSN (DSN=Nombre DSN)
sDsnName = "DSN=sqlsistema"
sDriver = "SQL Server"

' Modificamos el origen de datos de usuario especificado
dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

If dl = 1 Then
    MsgBox "El DSN de sistema se ha eliminado correctamente."
Else
    MsgBox "No se ha podido eliminar el DSN de sistema especificado."
End If


End Sub

Private Sub Command3_Click()
   rs.Open "SELECT dbo.notas_credito_fecha_mal_2.CLIENTE, dbo.notas_credito_fecha_mal_2.NOMBRE, dbo.notas_credito_fecha_mal_2.SERIE, dbo.notas_credito_fecha_mal_2.NUMERO_ANTERIOR, dbo.notas_credito_fecha_mal_2.NUMERO_ACTUAL, dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA FROM         dbo.notas_credito_fecha_mal_2 INNER JOIN dbo.TB_CLIENTES ON dbo.notas_credito_fecha_mal_2.CLIENTE = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID ", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "update tb_abono set vcha_abo_afectado_por = '" + CStr(rs!numero_actual) + "' where vcha_abo_clave_documento = '" + CStr(rs!numero_anterior) + "' and vcha_abo_tipo_documento = 'NCT' and vcha_abo_referencia = '" + rs!VCHA_CLI_REFERENCIA + "'", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   MsgBox "termino"
End Sub

Private Sub Command4_Click()
                      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
   'var_cadena = "SELECT     dbo.notas_credito_subir_oracle.NOTA, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO, dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AUD_MAQUINA FROM dbo.notas_credito_subir_oracle INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.notas_credito_subir_oracle.NOTA = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '02' AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = 'ncemx' INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID"
   var_cadena = "SELECT dbo.TB_CLIENTES.VCHA_CLI_REFERENCIA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AS nota, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA , dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE     (dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO in (7781)) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = 'ncemx') "
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_referencia_cliente_tienda = IIf(IsNull(rs!VCHA_CLI_REFERENCIA), "", rs!VCHA_CLI_REFERENCIA)
         var_total_neto = rs!floa_Car_importe_neto
         var_numero_nota = rs!nota
         var_dia_s = CStr(Day(rs!dtim_Car_fecha))
         var_mes_s = CStr(Month(rs!dtim_Car_fecha))
         var_año_s = CStr(Year(rs!dtim_Car_fecha))
         If Len(var_dia_s) = 1 Then
            var_dia_s = "0" + var_dia_s
         End If
         If Len(var_mes_s) = 1 Then
            var_mes_s = "0" + var_mes_s
         End If
         If Len(var_año_s) = 2 Then
            var_año_s = "20" + var_año_s
         End If
         var_fecha_s = "to_date('" + var_dia_s + "-" + var_mes_s + "-" + var_año_s + "','DD-MM-YYYY')"
         
         rsaux8.Open "CALL SP_AGREGA_ABONO('" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(var_total_neto)) + "," + CStr(CDbl(var_total_neto)) + "," + var_fecha_s + "," + var_fecha_s + ",'" + CStr(var_numero_nota) + "','','NCT','SUBIDAS NOTAS MAL')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
         
         rs.MoveNext
   Wend
End Sub

Private Sub Command5_Click()
    If cnn_clientes_tiendas.State = 0 Then
       cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
    End If
    rs.Open "select a.*, vcha_cli_referencia as vcha_cli_referencia from notas_credito_fecha_mal_2 a, tb_clientes b where cliente = vcha_cli_clave_id ", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          var_referencia = Trim(rs!VCHA_CLI_REFERENCIA)
          If var_referencia = "010700004897" Then
             MsgBox "ya"
             var_i = var_i
          End If
          VAR_NOTAS = "('" + CStr(rs!numero_anterior) + "','" + CStr(rs!numero_actual) + "')"
          rsaux.Open "select * from tb_abono where vcha_abo_referencia = '" + CStr(var_referencia) + "' and vcha_abo_clave_documento in " + VAR_NOTAS, cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
          'MsgBox "select * from tb_abono where vcha_abo_referencia = '" + CStr(var_referencia) + "' and vcha_abo_clave_documento in " + var_notas
          var_i = 0
          While Not rsaux.EOF
             var_i = var_i + 1
             rsaux.MoveNext
          Wend
          If var_i > 1 Then
             MsgBox var_referencia + " " + VAR_NOTAS
          End If
          rsaux.Close
          rs.MoveNext
    Wend
    rs.Close
    
End Sub

Private Sub Command6_Click()
     Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
     Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I

    rsaux10.Open "select * from tb_encabezado_cartera where vcha_Car_tipo_documento = 'NC' AND vcha_Ser_Serie_id = 'ncemyg' and inte_car_numero = 360", cnn, adOpenDynamic, adLockOptimistic
    var_i = 0
    While Not rsaux10.EOF
          var_serie = "SUMYG"
          var_cadena = "insert into tb_Encabezado_cartera "
          var_cadena = var_cadena + " (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_Ser_Serie_id, vcha_Car_tipo_documento, vcha_Car_documento, vcha_car_clase_id, inte_Car_numero, char_Car_afectacion, dtim_car_fecha,"
          var_cadena = var_cadena + " vcha_age_agente_id, vcha_gac_grupo_actual_id, vcha_gre_grupo_real_id, vcha_tit_titular_id, vcha_cli_clave_id, inte_car_plazo, floa_car_porcentaje_iva, floa_Car_porcentaje_impuesto_1, floa_Car_porcentaje_impuesto_2, floa_car_porcentaje_descuento_1, floa_Car_porcentaje_descuento_2, floa_car_porcentaje_descuento_3, floa_car_importe_total, floa_car_importe_iva, floa_Car_importe_impuesto_1, floa_car_importe_impuesto_2, floa_car_importe_descuento_1, floa_car_importe_descuento_2, floa_car_importe_descuento_3, floa_Car_subimporte, floa_car_importe_neto, vcha_aud_usuario, vcha_aud_maquina, vcha_aud_fecha, vcha_mon_moneda_id, floa_car_tipo_cambio, char_car_estatus, inte_Car_nota_credito_aplicada) values "
          var_cadena = var_cadena + " ('" + rsaux10!VCHA_EMP_EMPRESA_ID + "', '" + rsaux10!VCHA_UOR_UNIDAD_ID + "', '" + var_serie + "', 'SU', 'SU', 'SU', " + CStr(rsaux10!inte_Car_numero) + ", '+', GETDATE(),"
          var_cadena = var_cadena + " '" + rsaux10!VCHA_AGE_AGENTE_ID + "', '" + rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rsaux10!vcha_gre_grupo_real_id + "', '" + rsaux10!vcha_tit_titular_id + "', '" + rsaux10!vcha_cli_clave_id + "', 0, 16, 0, 0, 0, 0, 0, " + CStr(rsaux10!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rsaux10!floa_car_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(rsaux10!floa_car_subimporte) + ", " + CStr(rsaux10!floa_Car_importe_neto) + ", 'FERNANDO SERNA', 'FSERNAPORT', GETDATE(), '" + rsaux10!vcha_mon_moneda_id + "', " + CStr(rsaux10!floa_car_tipo_cambio) + ", 'I', 1)"
          rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
          rsaux.Open "INSERT INTO TB_eSTADO_CUENTA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ECU_MOVIMIENTO_CARGO, VCHA_ECU_SERIE_CARGO, INTE_ECU_NUMERO_cARGO, FLOA_ECU_IMPORTE_CARGO) values ('" + rsaux10!VCHA_EMP_EMPRESA_ID + "','" + rsaux10!VCHA_UOR_UNIDAD_ID + "','SU','" + var_serie + "'," + CStr(rsaux10!inte_Car_numero) + "," + CStr(rsaux10!floa_Car_importe_neto) + ")", cnn, adOpenDynamic, adLockOptimistic
          
          
          rs.Open "SELECT * FROM TB_sERIES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
          var_numero_folio = rs!inte_ser_nota_credito
          var_serie = rs!vcha_Ser_Serie_id
          rs.Close
          var_insertar = False
          'var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", "BO", "BO", var_numero_folio, "-", "", "", 0, CStr(Date), rsuax10!vcha_age_Agente_id, rsaux10!vcha_gac_grupo_actual_id, rsaux10!vcha_gre_grupo_real_id, rsaux10!vcha_tit_titular_id, rsaux10!vcha_cli_clave_id, "", 0, 16, 0, 0, 0, 0, 0, rsaux10!floa_car_importe_total, rsaux10!floa_car_importe_iva, 0, 0, 0, 0, 0, rsaux10!floa_car_subimporte, rsaux10!floa_car_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "1", 1, var_serie, "")
          
          
          var_cadena = "insert into tb_Encabezado_cartera "
          var_cadena = var_cadena + " (vcha_Emp_empresa_id, vcha_uor_unidad_id, vcha_Ser_Serie_id, vcha_Car_tipo_documento, vcha_Car_documento, vcha_car_clase_id, inte_Car_numero, char_Car_afectacion, dtim_car_fecha,"
          var_cadena = var_cadena + " vcha_age_agente_id, vcha_gac_grupo_actual_id, vcha_gre_grupo_real_id, vcha_tit_titular_id, vcha_cli_clave_id, inte_car_plazo, floa_car_porcentaje_iva, floa_Car_porcentaje_impuesto_1, floa_Car_porcentaje_impuesto_2, floa_car_porcentaje_descuento_1, floa_Car_porcentaje_descuento_2, floa_car_porcentaje_descuento_3, floa_car_importe_total, floa_car_importe_iva, floa_Car_importe_impuesto_1, floa_car_importe_impuesto_2, floa_car_importe_descuento_1, floa_car_importe_descuento_2, floa_car_importe_descuento_3, floa_Car_subimporte, floa_car_importe_neto, vcha_aud_usuario, vcha_aud_maquina, vcha_aud_fecha, vcha_mon_moneda_id, floa_car_tipo_cambio, char_car_estatus, inte_Car_nota_credito_aplicada) values "
          var_cadena = var_cadena + " ('" + rsaux10!VCHA_EMP_EMPRESA_ID + "', '" + rsaux10!VCHA_UOR_UNIDAD_ID + "', '" + var_serie + "', 'NC', 'BO', 'BO', " + CStr(var_numero_folio) + ", '-', GETDATE(),"
          var_cadena = var_cadena + " '" + rsaux10!VCHA_AGE_AGENTE_ID + "', '" + rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rsaux10!vcha_gre_grupo_real_id + "', '" + rsaux10!vcha_tit_titular_id + "', '" + rsaux10!vcha_cli_clave_id + "', 0, 16, 0, 0, 0, 0, 0, " + CStr(rsaux10!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rsaux10!floa_car_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(rsaux10!floa_car_subimporte) + ", " + CStr(rsaux10!floa_Car_importe_neto) + ", 'FERNANDO SERNA', 'FSERNAPORT', GETDATE(), '" + rsaux10!vcha_mon_moneda_id + "', " + CStr(rsaux10!floa_car_tipo_cambio) + ", 'I', 1)"
          rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
          'rsaux.Open "INSERT INTO TB_eSTADO_CUENTA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ECU_MOVIMIENTO_CARGO, VCHA_ECU_SERIE_CARGO, INTE_ECU_NUMERO_cARGO, FLOA_ECU_IMPORTE_CARGO) values ('" + rsaux10!VCHA_EMP_EMPRESA_ID + "','" + rsaux10!VCHA_UOR_UNIDAD_ID + "','" + rsaux10!vcha_car_documento + "','" + var_serie + "'," + CStr(rsaux10!inte_car_numero) + "," + CStr(rsaux10!floa_car_importe_neto) + ")", cnn, adOpenDynamic, adLockOptimistic
          
          
          
          
          rsaux3.Open "insert into tb_secuencia_notas_credito (vcha_emp_empresa_id, vcha_Ser_serie_id, inte_snc_numero_anterior, inte_snc_numero_actual) values ('" + var_empresa + "', '" + var_serie + "', " + CStr(var_numero_folio) + ", " + CStr(var_numero_folio) + ")", cnn, adOpenDynamic, adLockOptimistic
          
          
          rsaux3.Open "INSERT INTO TB_ESTADO_CUENTA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ECU_MOVIMIENTO_CARGO, VCHA_ECU_SERIE_CARGO, INTE_ECU_NUMERO_CARGO, FLOA_ECU_IMPORTE_CARGO, VCHA_ECU_MOVIMIENTO_ABONO, VCHA_ECU_SERIE_ABONO, INTE_ECU_NUMERO_ABONO, FLOA_ECU_IMPORTE_ABONO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "','SU','SUMYG'," + CStr(rsaux10!inte_Car_numero) + ",0, 'BO','" + var_serie + "'," + CStr(var_numero_folio) + "," + CStr(rsaux10!floa_Car_importe_neto) + ")", cnn, adOpenDynamic, adLockOptimistic
          rs.Open "update tb_series set inte_ser_nota_credito = inte_ser_nota_credito + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
          
   '''''  IMPRESION DE LA NOTA DE CARGO
          var_k = var_numero_nota_inicio
          rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BO' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             Open (App.Path & "\renombra" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
             Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
             Close #2
             Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi") For Output As #1
             'Open ("c:\NC_" + Trim(var_serie) + Trim(Str(rs!inte_car_numero)) + ".fi") For Output As #1
             var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + rs!vcha_Ser_Serie_id + Chr(13) + "folio=" + CStr(rs!inte_Car_numero) + Chr(13)
             var_año = CStr(Year(rs!dtim_Car_fecha))
             var_mes = CStr(Month(rs!dtim_Car_fecha))
             var_dia = CStr(Day(rs!dtim_Car_fecha))
             var_hora = CStr(Hour(rs!dtim_Car_fecha))
             var_minuto = CStr(Minute(rs!dtim_Car_fecha))
             var_segundo = CStr(Second(rs!dtim_Car_fecha))
             If Len(var_año) = 2 Then
                var_año = "20" + var_año
             End If
             If Len(var_mes) = 1 Then
                var_mes = "0" + var_mes
             End If
             If Len(var_dia) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(var_hora) = 1 Then
                var_hora = "0" + var_hora
             End If
             If Len(var_minuto) = 1 Then
                var_minuto = "0" + var_minuto
             End If
             If Len(var_segundo) = 1 Then
                var_segundo = "0" + var_segundo
             End If
                              
             var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
             var_rfc_cliente = ""
             If var_rfc_cliente_1 = "" Then
                var_rfc_cliente = "XAXX010101000"
             Else
                For var_j = 1 To Len(var_rfc_cliente_1)
                    If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                       If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                          If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                             var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                          End If
                       End If
                    End If
                Next var_j
             End If
                             
                              
             var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
             var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
             var_cadena = var_cadena + "noAprobacion=" + Chr(13)
             var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
             var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
             var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
             var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
             If var_rfc_cliente = "XAXX010101000" Then
                var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
             Else
                var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_car_subimporte / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
             End If
             var_cadena = var_cadena + "descuento=" + Chr(13)
             var_cadena = var_cadena + "descuento1=" + Chr(13)
             var_cadena = var_cadena + "descuento2=" + Chr(13)
             var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
             var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
             var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
             var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
             If rsaux2.State = 1 Then
                rsaux2.Close
             End If
             rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
             var_certificado = rsaux2!vcha_emp_certificado
             var_expedido = rsaux2!vcha_emp_expedido
             If var_rfc_cliente = "XAXX010101000" Then
                var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,##0.000000") + Chr(13)
             Else
                var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
             End If
             var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
             var_cadena = var_cadena + "retencion=" + Chr(13)
             var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
             var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "<Emisor>" + Chr(13)
             var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
             var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
             var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
             var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
             var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
             var_cadena = var_cadena + "enoInterior=" + Chr(13)
             var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
             var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
             var_cadena = var_cadena + "ereferencia=" + Chr(13)
             var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
             var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
             var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
             var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
             var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
             var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
             var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                             
             var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "ex_calle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
             var_cadena = var_cadena + "ex_noExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
             var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
             var_cadena = var_cadena + "ex_colonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
             var_cadena = var_cadena + "ex_localidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
             var_cadena = var_cadena + "ex_referencia=" + Chr(13)
             var_cadena = var_cadena + "ex_municipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
             var_cadena = var_cadena + "ex_estado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
             var_cadena = var_cadena + "ex_pais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
             var_cadena = var_cadena + "ex_codigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
             var_cadena = var_cadena + "</ExpedidoEn>"
                             
                             
                             
                             
             var_cadena = var_cadena + "<Receptor>" + Chr(13)
             var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
             rsaux2.Close
                                          
             var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
             var_rfc_cliente = ""
             If var_rfc_cliente_1 = "" Then
                var_rfc_cliente = "XAXX010101000"
             Else
                For var_j = 1 To Len(var_rfc_cliente_1)
                    If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                       If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                          If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                             var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                          End If
                       End If
                    End If
                Next var_j
             End If
             If var_empresa = "03" Or var_empresa = "28" Then
                var_rfc_cliente = "XEXX010101000"
             End If
             var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
             var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
             var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "<Cliente>" + Chr(13)
             var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
             var_cadena = var_cadena + "calle=" + Chr(13)
             var_cadena = var_cadena + "noExterior=" + Chr(13)
             var_cadena = var_cadena + "noInterior=" + Chr(13)
             var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
             var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + Chr(13)
             rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
             var_cadena = var_cadena + "referencia=" + Chr(13)
             var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
             var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
             VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
             If Trim(VAR_NOMBRE_PAIS) = "" Then
                VAR_NOMBRE_PAIS = "MEXICO"
             End If
             var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
             var_cadena = var_cadena + Chr(13)
             var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
             var_cadena = var_cadena + "tel=" + Chr(13)
             var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
             var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                             
             var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
             var_cadena = var_cadena + "endomicilio=" + Chr(13)
             var_cadena = var_cadena + "encalle=" + Chr(13)
             var_cadena = var_cadena + "ennoExterior=" + Chr(13)
             var_cadena = var_cadena + "ennoInterior=" + Chr(13)
             var_cadena = var_cadena + "encolonia=" + Chr(13)
             var_cadena = var_cadena + "enlocalidad=" + Chr(13)
             var_cadena = var_cadena + "enreferencia=" + Chr(13)
             var_cadena = var_cadena + "enmunicipio=" + Chr(13)
             var_cadena = var_cadena + "enestado=" + Chr(13)
             var_cadena = var_cadena + "enpais=" + Chr(13)
             var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
             var_cadena = var_cadena + "entel=" + Chr(13)
             var_cadena = var_cadena + "enemail=" + Chr(13)
             var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "<Concepto>" + Chr(13)
                             
                              
             var_i = 1
             pxx = CStr(var_i)
             If Len(pxx) = 1 Then
                pxx = "0" + pxx
             End If
             var_cadena = var_cadena + "p" + pxx + "_cantidad=1" + Chr(13)
             var_cadena = var_cadena + "p" + pxx + "_unidad=BO" + Chr(13)
             var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
             var_linea = txt_clase + Str(rs!inte_Car_numero) + " SUSTITUCION DE NOTA DE CREDITO NUMERO " + CStr(rsaux10!inte_Car_numero)
             var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
             If var_rfc_cliente = "XAXX010101000" Then
                var_importe_str = ((IIf(IsNull(rsaux10!floa_Car_importe_neto), 0, rsaux10!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)))
             Else
                var_importe_str = ((IIf(IsNull(rsaux10!floa_Car_importe_neto), 0, rsaux10!floa_Car_importe_neto)) / (1 + (16 / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)))
             End If
             var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
             var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
             'MsgBox var_cadena
             var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "<Otros>" + Chr(13)
             var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
             rs.MoveFirst
             var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
             var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
             rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
             var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
             rsaux1.Close
             var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
             var_cadena = var_cadena + "pedido=" + Chr(13)
             var_cadena = var_cadena + "Embarque=" + Chr(13)
             var_referencia_Bancaria = ""
             var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
             var_cadena = var_cadena + "fechaPedido=" + Chr(13)
             var_cadena = var_cadena + "expedicion=" + Chr(13)
             var_cadena = var_cadena + "observaciones=" + Chr(13)
             var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
             var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
             var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
             var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
             var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                              
             rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux11.EOF Then
                var_cadena = var_cadena + "agente=" + rsaux11!VCHA_AGE_AGENTE_ID + " " + rsaux11!VCHA_AGE_NOMBRE + Chr(13)
             End If
             rsaux11.Close
                             
             If var_empresa = "02" Or var_empresa = "15" Or var_empresa = "16" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
                var_cadena = var_cadena + "formato=MHNCVTH_V01.dat" + Chr(13)
             End If
             If var_empresa = "07" Then
                var_cadena = var_cadena + "formato=MHNCARE_V01.dat" + Chr(13)
             End If
             If var_empresa = "31" Then
                var_cadena = var_cadena + "formato=MHNCCAN_V01.dat" + Chr(13)
             End If
             If var_empresa = "42" Then
                var_cadena = var_cadena + "formato=MHNCCMA_V01.dat" + Chr(13)
             End If
             If var_empresa = "41" Then
                var_cadena = var_cadena + "formato=MHNCCOP_V01.dat" + Chr(13)
             End If
             If var_empresa = "150000" Then
                var_cadena = var_cadena + "formato=MHNCERE_V01.dat" + Chr(13)
             End If
             If var_empresa = "33" Then
                var_cadena = var_cadena + "formato=MHNCMPU_V01.dat" + Chr(13)
             End If
             If var_empresa = "34" Then
                var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
             End If
             If var_empresa = "160000" Then
                var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
             End If
             If var_empresa = "36" Then
                var_cadena = var_cadena + "formato=MHNCSME_V01.dat" + Chr(13)
             End If
             If var_empresa = "30" Then
                var_cadena = var_cadena + "formato=MHNCTUR_V01.dat" + Chr(13)
             End If
             If var_empresa = "44" Then
                var_cadena = var_cadena + "formato=MHNCUTV_V01.dat" + Chr(13)
             End If
             If var_empresa = "38" Then
                var_cadena = var_cadena + "formato=MHNCVIA_V01.dat" + Chr(13)
             End If
             If var_empresa = "40" Then
                var_cadena = var_cadena + "formato=MHNCVIN_V01.dat" + Chr(13)
             End If
             If var_empresa = "43" Then
                var_cadena = var_cadena + "formato=MHNCVOP_V01.dat" + Chr(13)
             End If
                              
                              
             var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "<addenda>" + Chr(13)
             var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
             var_cadena = var_cadena + "</Factura>"
             Print #1, var_cadena
             Close #1
                              
             var_Archivo = App.Path & "\renombra" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".bat"
             x = Shell(var_Archivo, vbHide)
          End If
          rs.Close
          rsaux10.MoveNext
    Wend
    rsaux10.Close
    MsgBox var_i
End Sub

Private Sub Command7_Click()
                           
                      cnnoracle.Close
                      cnnoracle.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=ap;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
                      cnnoracle.CursorLocation = adUseClient
                      rsaux9.Open "SELECT VCHA_EMO_REFERENCIA, DTIM_EMO_FECHA, inte_emo_numero FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_MOV_MOVIMIENTO_ID = 'EI' AND VCHA_EMP_EMPRESA_ID = '31' AND CHAR_EMO_ESTATUS = 'I' and dtim_emo_Fecha >= {d '2011-06-01'} and dtim_emo_fecha < {d '2011-06-20'} order by dtim_emo_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                      While Not rsaux9.EOF
                           'var_fecha = Date - 7
                           var_fecha = rsaux9!dtim_emo_fecha
                           var_numero_folio = rsaux9!inte_emo_numero
                           var_factura = IIf(IsNull(rsaux9!vcha_emo_Referencia), "", rsaux9!vcha_emo_Referencia)
                           var_cadena = "SELECT     SUM(FLOA_ENT_CANTIDAD) as cantidad , SUM(FLOA_ENT_CANTIDAD * FLOA_ENT_PRECIO) as precio,SUM(FLOA_ENT_CANTIDAD * FLOA_ENT_COSTO) as costo From dbo.TB_ENTRADAS WHERE     (VCHA_EMP_EMPRESA_ID = '31') AND (VCHA_UOR_UNIDAD_ID = '26') AND (VCHA_MOV_MOVIMIENTO_ID = 'EI') and inte_ent_numero = " + CStr(var_numero_folio) + " GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO"
                           rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           var_clave_movimiento = "EI"
                           If var_clave_movimiento = "EI" Then
                              rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "' AND POLIZA_ID = 2", cnnoracle, adOpenDynamic, adLockOptimistic
                              concepto_1 = "EI:" + CStr(var_numero_folio) + ";" + "Fac:" + var_factura
                              CONCEPTO_2 = "FACTURA NUM: " + var_factura
                              CONCEPTO_3 = "POLIZA FACTURAS INTERCOMPAÑIA"
                           End If
                           While Not rsaux11.EOF
                                 var_tipo_poliza = rsaux11!tipo
                                 var_origen_poliza = rsaux11!Origen
                                 var_categoria_poliza = rsaux11!categoria
                                 var_moneda_poliza = rsaux11!moneda
                                 var_segmento1_poliza = rsaux11!segmento1
                                 var_segmento2_poliza = rsaux11!segmento2
                                 var_segmento3_poliza = rsaux11!segmento3
                                 'If rsaux11!SEGMENTO4 = 2140 Then
                                 '   var_segmento4_poliza = "1161"
                                 'Else
                                    var_segmento4_poliza = rsaux11!SEGMENTO4
                                 'End If
                                 var_segmento5_poliza = rsaux11!segmento5
                                 var_segmento6_poliza = rsaux11!segmento6
                                 var_segmento7_poliza = rsaux11!segmento7
                                 var_juego_libros_poliza = rsaux11!juego_libros
                                 var_descripcion_poliza = rsaux11!descripcion
                                 var_cargo_poliza = rsaux11!cargo
                                 var_abono_poliza = rsaux11!abono
                                 var_precio = rsaux11!Precio
                                 If var_precio = 1 Then
                                    If rsaux10.EOF Then
                                       var_importe_precio = 0
                                    Else
                                       var_importe_precio = rsaux10!Precio
                                    End If
                                 Else
                                    If rsaux10.EOF Then
                                       var_importe_precio = 0
                                    Else
                                       var_importe_precio = IIf(IsNull(rsaux10!Costo), 0, rsaux10!Costo)
                                       'MsgBox CStr(var_importe_precio)
                                    End If
                                 End If
                                 var_cadena = "InsERT INTO IN_TB_POLIZAS_INT (STATUS, SET_OF_BOOKS_ID, USER_JE_SOURCE_NAME, USER_JE_CATEGORY_NAME, ACCOUNTING_DATE, CURRENCY_CODE, DATE_CREATED, ACTUAL_FLAG,  SEGMENT1, SEGMENT2, SEGMENT3, SEGMENT4, SEGMENT5, SEGMENT6, SEGMENT7, ENTERED_DR, ENTERED_CR, ACCOUNTED_DR, ACCOUNTED_CR, GROUP_ID, REFERENCE4, REFERENCE5, REFERENCe10, REFERENCE1, REFERENCE2, CREATED_BY)"
                                 If var_cargo_poliza = 1 Then
                                    var_cadena = var_cadena & " VALUES ('NEW', " & CStr(var_juego_libros_poliza) & ",'" & var_origen_poliza & "','" & var_categoria_poliza & "',TO_DATE('" & CStr(Format(var_fecha, "dd/mm/yyyy")) & "','DD/MM/YYYY'),'" & var_moneda_poliza & "',TO_DATE('" & CStr(Format(var_fecha, "dd/mm/yyyy")) & "','DD/MM/YYYY'),'A','" & var_segmento1_poliza & "','" & var_segmento2_poliza & "','" & var_segmento3_poliza & "','" & var_segmento4_poliza & "','" & var_segmento5_poliza & "','" & var_segmento6_poliza & "','" & var_segmento7_poliza & "'," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",0," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",0,1,'" & concepto_1 & "','" & CONCEPTO_2 & "','" & var_descripcion_poliza & "','" & CONCEPTO_3 & "','" & CONCEPTO_3 & "',1143)"
                                 Else
                                    var_cadena = var_cadena & " VALUES ('NEW', " & CStr(var_juego_libros_poliza) & ",'" & var_origen_poliza & "','" & var_categoria_poliza & "',TO_DATE('" & CStr(Format(var_fecha, "dd/mm/yyyy")) & "','DD/MM/YYYY'),'" & var_moneda_poliza & "',TO_DATE('" & CStr(Format(var_fecha, "dd/mm/yyyy")) & "','DD/MM/YYYY'),'A','" & var_segmento1_poliza & "','" & var_segmento2_poliza & "','" & var_segmento3_poliza & "','" & var_segmento4_poliza & "','" & var_segmento5_poliza & "','" & var_segmento6_poliza & "','" & var_segmento7_poliza & "',0," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",0," & CStr(IIf(IsNull(var_importe_precio), 0, var_importe_precio)) & ",1,'" & concepto_1 & "','" & CONCEPTO_2 & "','" & var_descripcion_poliza & "','" & CONCEPTO_3 & "','" & CONCEPTO_3 & "',1143)"
                                 End If
                                 'MsgBox var_cadena
                                 rsaux7.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                                 rsaux11.MoveNext
                           Wend
                           rsaux11.Close
                           rsaux10.Close
                           rsaux9.MoveNext
                      Wend
                      rsaux9.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_linea = True
   lv_lineas.SmallIcons = ImageList
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   rs.Close
   txt_lineas(0).Enabled = False
   If var_clave_usuario_global = "U0000000182" Then
      Me.txt_lineas(0).Enabled = False
      Me.txt_lineas(1).Enabled = False
      Me.txt_lineas(2).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_linea = False
   End If
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub lv_lineas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lineas, ColumnHeader)
End Sub

Private Sub lv_lineas_ItemClick(ByVal item As MSComctlLib.ListItem)
   Set lv_lineas.selectedItem = item
   pro_textos
   var_modifica_registro_linea = True
   txt_lineas(0).Enabled = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_lineas.SetFocus
      Call pro_avanzar(Me, lv_lineas, Button)
      lv_lineas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_lineas.ListItems(1).Selected = True
      lv_lineas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_lineas = lv_lineas.ListItems.Count
      lv_lineas.ListItems(numero_items_lineas).Selected = True
      lv_lineas.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_lineas()
   Dim ok As Boolean
   Set TB_LINEAS = New TB_LINEAS
   Set TB_BITACORA_LINEAS = New TB_BITACORA_LIENAS
   ok = True
   If txt_lineas(0) <> "" And txt_lineas(1) <> "" Then
      If var_hubo_cambios Then
         rs.Open "SELECT * FROM TB_LINEAS WHERE VCHA_LIN_LINEA_ID = '" + txt_lineas(0) + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_LINEAS.Anadir(txt_lineas(0), txt_lineas(1), txt_lineas(2))
         If ok Then
            If IsNumeric(Me.txt_capacidad) Then
               rsaux.Open "update tb_lineas set floa_lin_capacidad = " + Me.txt_capacidad + " where vcha_lin_linea_id = '" + Me.txt_lineas(0) + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            bitacora = True
            If var_modifica_registro_linea = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_LINEAS.Anadir(txt_lineas(0), "VCHA_LIN_NOMBRE", var_operacion_bitacora, "", txt_lineas(1), var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_lineas(0) Then
                  bitacora = TB_BITACORA_LINEAS.Anadir(txt_lineas(0), "VCHA_LIN_LINEA_ID", var_operacion_bitacora, rs(0), txt_lineas(0), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_lineas(1) Then
                  bitacora = TB_BITACORA_LINEAS.Anadir(txt_lineas(0), "VCHA_LIN_NOMBRE", var_operacion_bitacora, rs(1), txt_lineas(1), var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_lineas(0).Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_lineas.ListItems.Count
            var_modifica_registro_linea = True
         Else
            MsgBox "No se puede grabar registro: " + TB_LINEAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_LINEAS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_lineas()
   Dim var_llave_usuarios As String
   On Error GoTo salir:
   Set TB_LINEAS = New TB_LINEAS
   Set TB_BITACORA_LINEAS = New TB_BITACORA_LIENAS
   ok = True
   If txt_lineas(0) <> "" And txt_lineas(1) <> "" And var_modifica_registro_linea = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_LINEAS.Eliminar(txt_lineas(0))
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_LINEAS.Anadir(txt_lineas(0), "VCHA_LIN_NOMBRE", var_operacion_bitacora, txt_lineas(1), "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_lineas = numero_items_lineas - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_lineas.ListItems.Remove (lv_lineas.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_lineas.ListItems.Count
         lv_lineas.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede eliminar registro: " + TB_LINEAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_LINEAS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_lineas", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_lineas.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      'list_item.SubItems(3) = IIf(IsNull(rs!floa_lin_capacidad), "", rs!floa_lin_capacidad)
      rs.MoveNext:
      numero_items_lineas = numero_items_lineas + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_lineas.ListItems.Count
   If var_n > 0 Then
      txt_lineas(0) = lv_lineas.selectedItem
      txt_lineas(1) = lv_lineas.selectedItem.SubItems(1)
      txt_lineas(2) = lv_lineas.selectedItem.SubItems(2)
      Me.txt_capacidad = lv_lineas.selectedItem.SubItems(3)
   End If
   var_numero_renglones = lv_lineas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_lineas.ColumnHeaders(2).Width = 3850
   Else
      lv_lineas.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_linea = True
   txt_lineas(0).Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_linea = False Then
       Set list_item = lv_lineas.ListItems.Add(, , txt_lineas(0))
       list_item.SubItems(1) = txt_lineas(1)
       list_item.SubItems(2) = txt_lineas(2)
       list_item.SubItems(3) = Me.txt_capacidad
       list_item.EnsureVisible
       list_item.Selected = True
       numero_items_lineas = numero_items_lineas + 1
    Else
       lv_lineas.ListItems.item(lv_lineas.selectedItem.Index).Checked = False
       lv_lineas.ListItems.item(lv_lineas.selectedItem.Index) = txt_lineas(0)
       lv_lineas.ListItems.item(lv_lineas.selectedItem.Index).ListSubItems(1) = txt_lineas(1)
       lv_lineas.ListItems.item(lv_lineas.selectedItem.Index).ListSubItems(2) = txt_lineas(2)
       lv_lineas.ListItems.item(lv_lineas.selectedItem.Index).ListSubItems(3) = Me.txt_capacidad
       lv_lineas.ListItems.item(lv_lineas.selectedItem.Index).Selected = True
    End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_lineas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_capacidad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_capacidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_lineas_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_lineas_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Index < 2 Then
         Call pro_enfoque(KeyAscii)
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      var_hubo_cambios = True
   End If
   If Index = 2 Then
      Select Case KeyAscii
      Case 48 To 57, 52, 13, 8, 46
      Case Else
          KeyAscii = 0
      End Select
      If KeyAscii = 13 Then
         If Me.cmd_guardar.Enabled = True Then
            Me.cmd_guardar.SetFocus
         End If
      End If
   End If
End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub
