Attribute VB_Name = "modRecSetToCombo"
Option Explicit
'modRecsetToCombo
'Coded by Legrev3@aol.com
'Populates a combo box 10 times faster for recordsets having 12,000+ records
'May 1, 2001

'**  Function Declarations:
#If Win32 Then
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
'**  Constant Definitions:
Private Const CB_ADDSTRING& = &H143
Private Const CB_RESETCONTENT& = &H14B
#End If 'WIN32

Public Function RecsetToCombo(hwnd As Long, rsRecSet As ADODB.Recordset, intCol As Integer) As Boolean
    Dim lngRetVal As Long
        
    'On Error GoTo LocalErrHandler:
    If rsRecSet.BOF And rsRecSet.EOF Then
       ' MsgBox "No Hay Registros Dados de Alta...", vbInformation + vbOKOnly
        RecsetToCombo = False
        Exit Function
    End If
    
    rsRecSet.MoveFirst
    Call SendMessageBynum(hwnd, CB_RESETCONTENT, 0, 0)
    
    Do Until rsRecSet.EOF
        If Not IsNull(rsRecSet(intCol).Value) And rsRecSet(intCol).Value <> var_nombre_planta Then
            lngRetVal = SendMessageByString(hwnd, CB_ADDSTRING, 0, rsRecSet(intCol).Value)
            'check value of lngRetVal here for errors if desired
        End If
        rsRecSet.MoveNext
    Loop
    RecsetToCombo = True
    Exit Function
LocalErrHandler:
    MsgBox "Error in filling combo box: " & vbCrLf & _
        Err.Number & "  " & Err.Description, vbCritical + vbOKOnly
    Err.Clear
    RecsetToCombo = False
End Function


