VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmcierre 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmd_cierre 
      Caption         =   "&Cierre Mensual"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmcierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim x As New cl_sql_manejo
Dim fReturn As Boolean
Option Explicit

Private Sub cmdAddColumn_Click()
End Sub



Private Sub cmdCreateTable_Click()

End Sub

Private Sub cmd_cierre_Click()
On Error Resume Next
Dim var_fecha As String
Dim i As Long, n As Long
    n = 0
    var_fecha = Replace(Date, "/", "")

    fReturn = x.CreateTable("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_ARTICULO_I")
    ' Primary key needs to be different for each table created
    If fReturn = False Then Exit Sub
    ' Add multiple columns with different datatypes to the table
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_DESCRIPCION", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_PRO_PROVEEDOR_ID", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_LINEA", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_SUBLINEA", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_UBICACION", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "BINT_ART_ULTCOSTO", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "BINT_ART_COSPROMEDIO", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_UNI_UNIDAD_ID", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_TALLA", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "MON_ART_MATPRIMA", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "MON_ART_AVIOS", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "MON_ART_MANOBRA", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "MON_ART_GASFABRICACION", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "MON_ART_PRELISTA", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "MON_ART_PRECOSTO", "money")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_TIPO", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "FLOA_ART_EXISTENCIA", "float")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ALM_ALMACEN_ID", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_PRORRATEAR", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_ART_STATUS", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "DTIM_AUD_FECHA", "datetime")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_AUD_USUARIO", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "VCHA_AUD_MAQUINA", "varchar", "50")
    fReturn = x.CreateColumn("vianney", "TB_ARTICULOS_" & var_fecha, "BINT_PLA_PLANTA_ID", "bigint")
    
    
    
    If fReturn = True Then
    
        rs.Open "select * from TB_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
        rsaux.Open "select * from TB_ARTICULOS_" & var_fecha, cnn, adOpenDynamic, adLockOptimistic
        While Not rs.EOF
            For i = 0 To 24
                rsaux.AddNew
                rsaux(i).Value = IIf(IsNull(rs(i).Value), "", rs(i).Value)
                rsaux.Update
            Next i
            rs.MoveNext
            If n = 100 Then n = 0
            PB1.Value = n
            n = n + 1
        Wend
        PB1.Value = 100
        MsgBox "Cierre Creado Exitosamente... Con Nombre Historico de : TB_ARTICULOS_" & var_fecha
    Else
        MsgBox "No se Pudo Crear Cierres"
    End If

End Sub


Sub xcc()
End Sub

Private Sub Form_Load()

End Sub
