Attribute VB_Name = "mod_despliega_lista"
Option Explicit
Public Execute As Boolean
Public Mytextbox As TextBox
Public Sub PopulateList(pList As ListView, pRst As ADODB.Recordset)
    On Error Resume Next
    Dim i As Integer, iColCount As Integer
    Dim sColName As String
    Dim sColValue As String
    Dim oCH As ColumnHeader
    Dim oLI As ListItem
    Dim oSI As ListSubItem
    Dim oFld As ADODB.Field

    With pList
        .View = lvwReport
        pRst.MoveFirst
        For Each oFld In pRst.Fields
            sColName = CkNuL(oFld.Name)
            Set oCH = .ColumnHeaders.Add()
            oCH.Text = sColName
            iColCount = iColCount + 1
        
        Next oFld

        While Not pRst.EOF
            i = 0
           
            sColValue = CkNuL(pRst.Fields(i).Value)
            Set oLI = .ListItems.Add()
            oLI.Text = sColValue
                        
            For i = 1 To iColCount
                Set oSI = oLI.ListSubItems.Add()
                oSI.Text = CkNuL(pRst(i))
            Next
            pRst.MoveNext
        Wend ' next record
    pRst.Close
    Set pRst = Nothing
    End With
End Sub

Private Function CkNuL(pVal As String) As String
    If IsMissing(pVal) Then
        CkNuL = ""
    ElseIf IsNull(pVal) Then
        CkNuL = ""
    Else
        CkNuL = Format(pVal)
    End If
End Function

Public Sub Autocomplete(Lvw As ListView, sFind, Mytextbox As TextBox)
Dim Lvfindtm As ListItem
Dim TempSelStart As Integer
Dim strTemp As String

Set Lvfindtm = Lvw.FindItem(sFind, lvwText, , lvwPartial)
If Not Lvfindtm Is Nothing Then
Lvfindtm.EnsureVisible
Lvfindtm.Selected = True

If Execute Then
TempSelStart = Mytextbox.SelStart
Mytextbox.Text = CStr(Lvfindtm)
If Not Mytextbox.Text = "" Then
Mytextbox.SelStart = TempSelStart
Mytextbox.SelLength = Len(Mytextbox.Text) - TempSelStart
    End If
        End If
            End If
End Sub


