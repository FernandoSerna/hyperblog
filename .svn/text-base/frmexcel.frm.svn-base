VERSION 5.00
Begin VB.Form frmexcel 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Excel"
      Height          =   435
      Left            =   1560
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "SubItem"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Text"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find Item"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "StringToFind"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox lv 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TIT!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   2985
   End
End
Attribute VB_Name = "frmexcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itmFound As ListItem
Const BadIP = "192.168.0.2"

Private Sub Command2_Click()
    Call SendToExcel(lv)
End Sub

Private Sub Form_Load()
    Dim clmX As ColumnHeader
    Dim listX As ListItem
    
    Set clmX = lv.ColumnHeaders.Add(, , "UserName", lv.Width / 4)
    Set clmX = lv.ColumnHeaders.Add(, , "UserID", lv.Width / 4)
    Set clmX = lv.ColumnHeaders.Add(, , "UserIP", lv.Width / 4)
    Set clmX = lv.ColumnHeaders.Add(, , "Last Action", lv.Width / 4)
    
    Set listX = lv.ListItems.Add(, , "Vegetto")
    lv.ListItems(1).SubItems(1) = "1-123"
    lv.ListItems(1).SubItems(2) = "127.0.0.1"

    Set listX = lv.ListItems.Add(, , "Gogeta")
    lv.ListItems(2).SubItems(1) = "0-123"
    lv.ListItems(2).SubItems(2) = "127.0.0.2"
    
    Set listX = lv.ListItems.Add(, , "Gotenks")
    lv.ListItems(3).SubItems(1) = "0-124"
    lv.ListItems(3).SubItems(2) = "127.0.0.3"
    
    Set listX = lv.ListItems.Add(, , "Billy")
    lv.ListItems(4).SubItems(1) = "1-124"
    lv.ListItems(4).SubItems(2) = "192.168.0.2"
    
    lv.BorderStyle = ccFixedSingle
    lv.View = lvwReport
End Sub

Private Sub Command1_Click()
    FindThingy Text1
End Sub

Private Sub lv_LostFocus()
    Dim i As Integer
    For i = 1 To lv.ListItems.Count
       lv.ListItems.Item(i).Selected = False
    Next i
End Sub

Public Sub FindThingy(ByVal StringToFind As String)
    ' option1, text, is for only searching column one
    ' option2, subitem, is for searching columns 2 and on
    
    ' now on what you wanted to do, you must put the UserName on column one
    ' this is how it refers to the list item, and it seems that is what you want to use for the reference

    If Option2.Value = True Then
        Set itmFound = lv.FindItem(StringToFind, lvwSubItem, , lvwPartial)
    ElseIf Option1.Value = True Then
        Set itmFound = lv.FindItem(StringToFind, lvwText, , lvwPartial)
    End If
    
    If itmFound Is Nothing Then
        MsgBox "No match found"
        Exit Sub
    Else
        itmFound.EnsureVisible
        itmFound.Selected = True
        lv.SetFocus
        
        'your ip thing
        'searching by username or userid
        'if billy is not found
        If Not itmFound = "Billy" Then
            ' then exit sub
            Exit Sub
        Else
            ' else, procede with the ip analysis
            ' find the rownumber
            RowNumber = lv.SelectedItem.Index
            ' if the ip column in that rownumber is the bad ip then
            If lv.ListItems(RowNumber).SubItems(2) = BadIP Then
                ' tell the admin that billy is bad
                MsgBox "Yeah, Billy is a bad person."
                lv.ListItems(RowNumber).SubItems(3) = "Billy did something"
                ' put the allow part here
            Else
                ' else tell the admin that billy is god
                MsgBox "Billy is clean... for now."
                ' put the disallow part here
            End If
        End If
    End If
End Sub

