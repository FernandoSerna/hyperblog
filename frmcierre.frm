VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmcierre 
   Caption         =   "Cierre Mensual de Inventario"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "frmcierre.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin vbskfree.Skinner Skinner1 
      Left            =   240
      Top             =   360
      _ExtentX        =   1270
      _ExtentY        =   1270
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSComctlLib.ProgressBar PB_1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdcierre 
      Caption         =   "&Cierre Mensual"
      Height          =   735
      Left            =   1920
      Picture         =   "frmcierre.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Mantiene el Estado del Inventario a la Fecha de Cierre"
      Height          =   915
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2520
   End
End
Attribute VB_Name = "frmcierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function fun_cierre() As Boolean
Dim var_fechatemp As Date
'On Error GoTo HELL:
    var_fecha_temp = Format(Date, "dd/mm/yyyy")
    
    rsaux.Open "select * from TB_CIERRE where TB_CIERRE.VCHA_CIE_FECHA = '" & var_fecha_temp & "'", cnn, adOpenDynamic, adLockOptimistic
    If rsaux.RecordCount = 0 Then '"& ddd &"'"
    rsaux.Close
        rs.Open "select * from TB_ARTICULOS", cnn, adOpenDynamic, adLockOptimistic
        rsaux.Open "select * from TB_CIERRE", cnn, adOpenDynamic, adLockOptimistic
        
        While Not rs.EOF
            rsaux.AddNew
            rsaux(0).Value = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            rsaux(1).Value = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rsaux(2).Value = IIf(IsNull(rs(2).Value), "", rs(2).Value)
            rsaux(3).Value = IIf(IsNull(rs(3).Value), "", rs(3).Value)
            rsaux(4).Value = IIf(IsNull(rs(4).Value), "", rs(4).Value)
            rsaux(5).Value = IIf(IsNull(rs(5).Value), "", rs(5).Value)
            rsaux(6).Value = IIf(IsNull(rs(6).Value), 0, rs(6).Value)
            rsaux(7).Value = IIf(IsNull(rs(7).Value), 0, rs(7).Value)
            rsaux(8).Value = IIf(IsNull(rs(8).Value), "", rs(8).Value)
            rsaux(9).Value = IIf(IsNull(rs(9).Value), "", rs(9).Value)
            rsaux(10).Value = IIf(IsNull(rs(10).Value), 0, rs(10).Value)
            rsaux(11).Value = IIf(IsNull(rs(11).Value), 0, rs(11).Value)
            rsaux(12).Value = IIf(IsNull(rs(12).Value), 0, rs(12).Value)
            rsaux(13).Value = IIf(IsNull(rs(13).Value), 0, rs(13).Value)
            rsaux(14).Value = IIf(IsNull(rs(14).Value), 0, rs(14).Value)
            rsaux(15).Value = IIf(IsNull(rs(15).Value), 0, rs(15).Value)
            rsaux(16).Value = IIf(IsNull(rs(16).Value), "", rs(16).Value)
            rsaux(17).Value = IIf(IsNull(rs(17).Value), "", rs(17).Value)
            rsaux(18).Value = IIf(IsNull(rs(18).Value), "", rs(18).Value)
            rsaux(19).Value = IIf(IsNull(rs(19).Value), "", rs(19).Value)
            rsaux(20).Value = IIf(IsNull(rs(20).Value), "", rs(20).Value)
            rsaux(21).Value = IIf(IsNull(rs(21).Value), "", rs(21).Value)
            rsaux(22).Value = IIf(IsNull(rs(22).Value), "", rs(22).Value)
            rsaux(23).Value = IIf(IsNull(rs(23).Value), "", rs(23).Value)
            rsaux(24).Value = IIf(IsNull(rs(24).Value), "", rs(24).Value)
            rsaux(25).Value = Format(Date, "dd/mm/yyyy")
            rsaux.Update
            rs.MoveNext
            If n = 100 Then n = 0
            PB_1.Value = n
            n = n + 1
        Wend
        PB_1.Value = 100
        fun_cierre = True
        rs.Close: rsaux.Close
    Else
        fun_cierre = False
        rsaux.Close
    End If
Exit Function
HELL:
    fun_cierre = False

    
End Function

Private Sub cmdcierre_Click()
Dim var_ultimo_dia_mes  As Date
Set clsdate = New clsdate
'    If Date = clsdate.LastOfMonth(Date) Then
        If fun_cierre Then
            SetTimer hwnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
            MsgBox "Se hizo el Cierre Existosamente .", , "TRANSACCIONES [ AVISO ]"
            Unload Me
        Else
            SetTimer hwnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
            MsgBox "Ya Esta Cerrado el Inventario a esta Fecha.", , "TRANSACCIONES [ AVISO ]"
            Unload Me
        End If
'    Else
'    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
'    MsgBox "Solo se puede Hacer Cierre el dia Ultimo del Mes.", vbCritical, "TRANSACCIONES [ AVISO ]"
'    End If
Set clsdate = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call menuvisible(Frmmenu2, True)
End Sub
