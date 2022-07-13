VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_ubicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta y modificación de ubicaciones"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1455
      TabIndex        =   27
      Top             =   135
      Width           =   5970
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1950
         Left            =   45
         TabIndex        =   28
         Top             =   420
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3440
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
            Text            =   "Clave"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         Caption         =   " Almacenes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Migrar "
      Height          =   300
      Left            =   675
      TabIndex        =   26
      Top             =   15
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Height          =   1485
      Left            =   90
      TabIndex        =   18
      Top             =   390
      Width           =   8070
      Begin VB.TextBox txt_codigo 
         Height          =   390
         Left            =   840
         TabIndex        =   24
         Top             =   930
         Width           =   1290
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   390
         Left            =   2145
         TabIndex        =   23
         Top             =   930
         Width           =   5850
      End
      Begin VB.TextBox txt_clave_almacen 
         Height          =   390
         Left            =   840
         TabIndex        =   20
         Top             =   495
         Width           =   1290
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   390
         Left            =   2145
         TabIndex        =   19
         Top             =   495
         Width           =   5850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   1005
         Width           =   600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Almacén:"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   570
         Width           =   660
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Caption         =   "  Datos "
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   7995
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3990
      Left            =   90
      TabIndex        =   10
      Top             =   1845
      Width           =   8070
      Begin VB.TextBox txt_ubicacion_6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2895
         TabIndex        =   8
         Top             =   3375
         Width           =   3645
      End
      Begin VB.TextBox txt_ubicacion_5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2895
         TabIndex        =   7
         Top             =   2820
         Width           =   3645
      End
      Begin VB.TextBox txt_ubicacion_4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2895
         TabIndex        =   6
         Top             =   2265
         Width           =   3645
      End
      Begin VB.TextBox txt_ubicacion_3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2895
         TabIndex        =   5
         Top             =   1695
         Width           =   3645
      End
      Begin VB.TextBox txt_ubicacion_2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2895
         TabIndex        =   4
         Top             =   1140
         Width           =   3645
      End
      Begin VB.TextBox txt_ubicacion_1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2895
         TabIndex        =   3
         Top             =   585
         Width           =   3645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 6:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   17
         Top             =   3465
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 5:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   16
         Top             =   2910
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   15
         Top             =   2355
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   14
         Top             =   1785
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   13
         Top             =   1230
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   12
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   " Ubicaciones"
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   30
         TabIndex        =   11
         Top             =   135
         Width           =   7995
      End
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   330
      Picture         =   "frmoracle_ubicaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_ubicaciones.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo "
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7845
      Picture         =   "frmoracle_ubicaciones.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   9
      Top             =   270
      Width           =   8220
   End
End
Attribute VB_Name = "frmoracle_ubicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_guardar_Click()
   If Me.txt_clave_almacen <> "" Then
      If Me.txt_codigo <> "" Then
         var_si = MsgBox("Confirmar la aplicación de los cambio", vbYesNo, "ATENCION")
         If var_si = 6 Then
            'rs.Open "UPDATE mtl_system_items_b SET  attribute2 ='" + Me.txt_ubicacion_1 + "', attribute3 ='" + Me.txt_ubicacion_2 + "', attribute4 = '" + Me.txt_ubicacion_3 + "', attribute5 = '" + Me.txt_ubicacion_4 + "', attribute6 = '" + Me.txt_ubicacion_5 + "', attribute7 = '" + Me.txt_ubicacion_6 + "' where  segment1 = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If txt_clave_almacen = "CDI_ALMPT" Then
               strconsulta = "UPDATE mtl_system_items_b SET  attribute2 = ?, attribute3 = ?, attribute4 = ?, attribute5 = ?, attribute6 = ?, attribute7 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_1)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_2)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_3)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_4)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_5)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_6)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
            End If
            
            'If Me.txt_ubicacion_1 <> "" Then
            strconsulta = "UPDATE mtl_system_items_b SET  attribute2 = ?, attribute3 = ?, attribute4 = ?, attribute5 = ?, attribute6 = ?, attribute7 = ? where  segment1 = ? and organization_id = ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_1)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_2)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_3)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_4)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_5)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_6)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                 .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
               strconsulta = "select * from xxvia_Tb_ubicaciones where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 1)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  strconsulta = "UPDATE xxvia_Tb_ubicaciones SET UBICACION = ? where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_1)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 1)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               Else
                  strconsulta = "INSERT INTO xxvia_Tb_ubicaciones (UBICACION, ALMACEN, CODIGO, ORGANIZACION, NUMERO) VALUES (?, ?, ?, ?, ?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_1)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 1)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
               rsaux9.Close
            'End If
            
            'If Me.txt_ubicacion_2 <> "" Then
               strconsulta = "select * from xxvia_Tb_ubicaciones where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 2)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  strconsulta = "UPDATE xxvia_Tb_ubicaciones SET UBICACION = ? where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_2)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 2)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               Else
                  strconsulta = "INSERT INTO xxvia_Tb_ubicaciones (UBICACION, ALMACEN, CODIGO, ORGANIZACION, NUMERO) VALUES (?, ?, ?, ?, ?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_2)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 2)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
               rsaux9.Close
            'End If
            
            'If Me.txt_ubicacion_3 <> "" Then
               strconsulta = "select * from xxvia_Tb_ubicaciones where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 3)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  strconsulta = "UPDATE xxvia_Tb_ubicaciones SET UBICACION = ? where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_3)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 3)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               Else
                  strconsulta = "INSERT INTO xxvia_Tb_ubicaciones (UBICACION, ALMACEN, CODIGO, ORGANIZACION, NUMERO) VALUES (?, ?, ?, ?, ?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_3)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 3)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
               rsaux9.Close
            'End If
            
            'If Me.txt_ubicacion_4 <> "" Then
               strconsulta = "select * from xxvia_Tb_ubicaciones where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 4)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  strconsulta = "UPDATE xxvia_Tb_ubicaciones SET UBICACION = ? where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_4)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 4)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               Else
                  strconsulta = "INSERT INTO xxvia_Tb_ubicaciones (UBICACION, ALMACEN, CODIGO, ORGANIZACION, NUMERO) VALUES (?, ?, ?, ?, ?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_4)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 4)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
               rsaux9.Close
            'End If
            
            'If Me.txt_ubicacion_5 <> "" Then
               strconsulta = "select * from xxvia_Tb_ubicaciones where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 5)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  strconsulta = "UPDATE xxvia_Tb_ubicaciones SET UBICACION = ? where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_5)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 5)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               Else
                  strconsulta = "INSERT INTO xxvia_Tb_ubicaciones (UBICACION, ALMACEN, CODIGO, ORGANIZACION, NUMERO) VALUES (?, ?, ?, ?, ?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_5)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 5)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
               rsaux9.Close
            'End If
            
            'If Me.txt_ubicacion_6 <> "" Then
               strconsulta = "select * from xxvia_Tb_ubicaciones where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 6)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  strconsulta = "UPDATE xxvia_Tb_ubicaciones SET UBICACION = ? where almacen = ? and codigo = ? and organizacion = ? and numero = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_6)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 6)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               Else
                  strconsulta = "INSERT INTO xxvia_Tb_ubicaciones (UBICACION, ALMACEN, CODIGO, ORGANIZACION, NUMERO) VALUES (?, ?, ?, ?, ?) "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ubicacion_6)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 6)
                       .Parameters.Append parametro
                  End With
                  Set rsaux10 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
               rsaux9.Close
            'End If
            
            MsgBox "Se han aplicado los cambios correctamente", vbOKOnly, "ATENCION"
         Else
            MsgBox "Código de artículo incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de almacén incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_ubicacion_1 = ""
   Me.txt_ubicacion_2 = ""
   Me.txt_ubicacion_3 = ""
   Me.txt_ubicacion_4 = ""
   Me.txt_ubicacion_5 = ""
   Me.txt_ubicacion_6 = ""
   If Me.txt_clave_almacen <> "" Then
      Me.txt_codigo.SetFocus
   Else
      Me.txt_clave_almacen.SetFocus
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   rs.Open "select segment1, nvl(attribute2,'') as ubicacion1, nvl(attribute3,'') as ubicacion2, nvl(attribute4,'') as ubicacion3, nvl(attribute5,'') as ubicacion4, nvl(attribute6,'') as ubicacion5, nvl(attribute7,'') as ubicacion6  from xxvia_system_items_b where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If IIf(IsNull(rs!UBICACION1), "", rs!UBICACION1) <> "" Then
            rsaux.Open "insert into xxvia_tb_ubicaciones (organizacion, almacen, codigo, numero, ubicacion) values (" + var_unidad_organizacional + ",'CDI_ALMPT','" + rs!SEGMENT1 + "',1,'" + rs!UBICACION1 + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If IIf(IsNull(rs!UBICACION2), "", rs!UBICACION2) <> "" Then
            rsaux.Open "insert into xxvia_tb_ubicaciones (organizacion, almacen, codigo, numero, ubicacion) values (" + var_unidad_organizacional + ",'CDI_ALMPT','" + rs!SEGMENT1 + "',2,'" + rs!UBICACION2 + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If IIf(IsNull(rs!UBICACION3), "", rs!UBICACION3) <> "" Then
            rsaux.Open "insert into xxvia_tb_ubicaciones (organizacion, almacen, codigo, numero, ubicacion) values (" + var_unidad_organizacional + ",'CDI_ALMPT','" + rs!SEGMENT1 + "',3,'" + rs!UBICACION3 + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If IIf(IsNull(rs!UBICACION4), "", rs!UBICACION4) <> "" Then
            rsaux.Open "insert into xxvia_tb_ubicaciones (organizacion, almacen, codigo, numero, ubicacion) values (" + var_unidad_organizacional + ",'CDI_ALMPT','" + rs!SEGMENT1 + "',4,'" + rs!UBICACION4 + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If IIf(IsNull(rs!UBICACION5), "", rs!UBICACION5) <> "" Then
            rsaux.Open "insert into xxvia_tb_ubicaciones (organizacion, almacen, codigo, numero, ubicacion) values (" + var_unidad_organizacional + ",'CDI_ALMPT','" + rs!SEGMENT1 + "',5,'" + rs!UBICACION5 + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If IIf(IsNull(rs!UBICACION6), "", rs!UBICACION6) <> "" Then
            rsaux.Open "insert into xxvia_tb_ubicaciones (organizacion, almacen, codigo, numero, ubicacion) values (" + var_unidad_organizacional + ",'CDI_ALMPT','" + rs!SEGMENT1 + "',6,'" + rs!UBICACION6 + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 1500
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_LostFocus()
   If Trim(txt_clave_almacen) <> "" Then
      rsaux.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + txt_clave_almacen + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         txt_nombre_almacen = rsaux!vcha_alm_nombre
         var_almacen_Destino = txt_almacen
      Else
         Me.txt_clave_almacen = ""
         Me.txt_nombre_almacen = ""
         MsgBox "El almacén no existe", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   Else
      MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_clave_Change()

End Sub

Private Sub txt_clave_LostFocus()

End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_clave_almacen = Me.lv_lista.selectedItem
      Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
      Me.txt_clave_almacen.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_clave_almacen.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_clave_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and attribute3 = 'INV_UBI'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre_almacen.SetFocus
   End If
End Sub

Private Sub txt_clave_almacen_LostFocus()
   If Trim(txt_clave_almacen) <> "" Then
      rsaux.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + txt_clave_almacen + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         txt_nombre_almacen = rsaux!vcha_alm_nombre
         var_almacen_Destino = txt_almacen
      Else
         Me.txt_clave_almacen = ""
         Me.txt_nombre_almacen = ""
         MsgBox "El almacén no existe", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   Else
      Me.txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_ubicacion_1 = ""
   Me.txt_ubicacion_2 = ""
   Me.txt_ubicacion_3 = ""
   Me.txt_ubicacion_4 = ""
   Me.txt_ubicacion_5 = ""
   Me.txt_ubicacion_6 = ""
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_articulos.Show 1
      Me.txt_codigo = var_codigo_busqueda
      Me.txt_descripcion = var_descripcion_busqueda
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Len(Trim(Me.txt_codigo)) = 5 And var_unidad_organizacional = 93 Then
      Me.txt_codigo = "000" + Me.txt_codigo
   End If
   If Trim(Me.txt_codigo) <> "" Then
      
      'rsaux9.Open "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo + "' order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select * from xxvia_system_items_b where  segment1 = ? and organization_id = ? "
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_unidad_organizacional)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         Me.txt_descripcion = IIf(IsNull(rsaux9!Description), "", rsaux9!Description)
         Me.txt_ubicacion_1 = ""
         Me.txt_ubicacion_2 = ""
         Me.txt_ubicacion_3 = ""
         Me.txt_ubicacion_4 = ""
         Me.txt_ubicacion_5 = ""
         Me.txt_ubicacion_6 = ""
      
         strconsulta = "select * from xxvia_Tb_ubicaciones where codigo = ? and organizacion = ?  and almacen = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_unidad_organizacional)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         While Not rsaux8.EOF
               If rsaux8!NUMERO = 1 Then
                  Me.txt_ubicacion_1 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 2 Then
                  Me.txt_ubicacion_2 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 3 Then
                  Me.txt_ubicacion_3 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 4 Then
                  Me.txt_ubicacion_4 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 5 Then
                  Me.txt_ubicacion_5 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 6 Then
                  Me.txt_ubicacion_6 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               rsaux8.MoveNext
         Wend
         rsaux8.Close
      Else
         MsgBox "El código del artículo no existe", vbOKOnly, "ATENCION"
      End If
      rsaux9.Close
   Else
      'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
      Me.txt_descripcion = ""
      Me.txt_ubicacion_1 = ""
      Me.txt_ubicacion_2 = ""
      Me.txt_ubicacion_3 = ""
      Me.txt_ubicacion_4 = ""
      Me.txt_ubicacion_5 = ""
      Me.txt_ubicacion_6 = ""
   End If
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_articulos.Show 1
      Me.txt_codigo = var_codigo_busqueda
      Me.txt_descripcion = var_descripcion_busqueda
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_1.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ubicacion_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_2.SetFocus
   End If
End Sub

Private Sub txt_ubicacion_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_3.SetFocus
   End If
End Sub

Private Sub txt_ubicacion_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_4.SetFocus
   End If
End Sub

Private Sub txt_ubicacion_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_5.SetFocus
   End If
End Sub

Private Sub txt_ubicacion_5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_6.SetFocus
   End If
End Sub

Private Sub txt_ubicacion_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
