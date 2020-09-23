VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "Borrowbook.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4680
      TabIndex        =   24
      Text            =   "Text4"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Text            =   "0"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4680
      TabIndex        =   17
      Text            =   "Text5"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   77070337
      CurrentDate     =   39707
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1800
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1800
      TabIndex        =   21
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1800
      TabIndex        =   20
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1800
      TabIndex        =   19
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Borrowed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Borrowed :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Borrower's Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3360
      Picture         =   "Borrowbook.frx":06EA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book NO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Borrower's ID :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   15
      Top             =   3120
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text = "" Then
    Combo1.Text = "1ST YR."
End If
End Sub

Private Sub Command1_Click()
If Text5.Text = "Blacklisted" Then
MsgBox "That User is Currently Blacklisted Change the Status first before you can let this Borrower Borrow"
ElseIf Text2.Text = "" Or Text1.Text = "" Then
    MsgBox "All fields required not to be a null value.", vbExclamation, "Library System"
    Exit Sub
    Else
  With Form1
    .Adodc3.Recordset.AddNew
    .Adodc3.Recordset.Fields(0) = (Label15.Caption)
    .Adodc3.Recordset.Fields(1) = (Text2.Text)
    .Adodc3.Recordset.Fields(2) = (Label13.Caption)
    .Adodc3.Recordset.Fields(3) = (Text1.Text)
    .Adodc3.Recordset.Fields(4) = (Label14.Caption)
    .Adodc3.Recordset.Fields(10) = "Borrowed"
    .Adodc3.Recordset.Fields(5) = Format(Label16.Caption, "mm/dd/yy")
    .Adodc3.Recordset.Fields(6) = Format(Text4.Text, "mm/dd/yy")
    Call SubtractBookQty
    .Adodc3.Recordset.Update
    Form1.Refresh
End With
MsgBox "New Barrowed Book/s has been added." & vbCrLf & "If you want to edit it, just delete it and 'Input again'.", vbInformation, "Library System"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Text2.Text = ""
'Text15.Caption = ""
Label14.Caption = ""
Text1.Text = ""
DTPicker1.Value = Format(Date, "mm/dd/yy")
'DTPicker2.Value = Format(Date, "mm/dd/yy")
End Sub
Private Sub Command3_Click()
Me.Hide
End Sub



Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
Label14.Caption = ""
    Text5.Text = ""
    Command1.Enabled = False
    Exit Sub
End If
Form1.Adodc1.Recordset.MoveFirst
Form1.Adodc1.Recordset.Find "BORROWERS_ID like '" & (Text1.Text) & "'"
If Form1.Adodc1.Recordset.EOF Then
    On Error Resume Next
    Label14.Caption = ""
    Text5.Text = ""
    DTPicker2.Enabled = False
    Command1.Enabled = False
    MsgBox "The Barrower's ID " & (DataCombo2.Text) & " Does not exist.Make sure it is correct.", vbExclamation, "Invalid Barrower's ID"
    DTPicker2.Value = Format(Date, "mm/dd/yy")
    Exit Sub

ElseIf Form1.Adodc1.Recordset.Fields(6) = "Blacklisted " Then
MsgBox "That User is Currently Blacklisted Change the Status first before you can let this Borrower Borrow"
Else
    Label14.Caption = (Form1.Adodc1.Recordset.Fields(1))
    Text5.Text = (Form1.Adodc1.Recordset.Fields(6))
    Command1.Enabled = True
End If
End Sub

Private Sub Form_Activate()
Text2.SetFocus
Text4.Text = Format(Date, "mm/dd/yy")
End Sub

Private Sub Form_Load()
Label16.Caption = Format(Date, "mm/dd/yy")
'DTPicker2.Value = Format(Date, "mm/dd/yy")
'Call BorrowedkBokNo
'Call BorrowedBorID
End Sub


Private Sub Text2_LostFocus()
If Text2.Text = "" Then
    Command2_Click
    Exit Sub
End If
Form1.Adodc2.Recordset.MoveFirst
Form1.Adodc2.Recordset.Find "ISBN like '" & (Text2.Text) & "'"
If Form1.Adodc2.Recordset.EOF Then
    On Error Resume Next
    MsgBox "The ISBN. " & (Text2.Text) & " Does not exist.Make sure it is correct.", vbExclamation, "Invalid ISBN."
    DataCombo1.SetFocus
    Command2_Click
    Exit Sub
Else
On Error Resume Next
    If Val(Form1.Adodc2.Recordset.Fields(9)) <= 0 Then
        MsgBox "All books in titled of " & (Form1.Adodc2.Recordset.Fields(2)) & " has all Borrowed.", vbExclamation, "Library System"
        Exit Sub
    End If
    Label13.Caption = (Form1.Adodc2.Recordset.Fields(2))
    Label15.Caption = (Form1.Adodc2.Recordset.Fields(0))
    DTPicker2.Enabled = True
End If
End Sub
Private Sub Text3_Change()
If Text3.Text = "" Then
Label7.ForeColor = vbRed
Command1.Enabled = False
Exit Sub
ElseIf Not (IsNumeric(Text3.Text)) Then
MsgBox "Input only a numeric value on the field Days Borrowed"
Command1.Enabled = False
Label7.ForeColor = vbRed
Else
Text4.Text = Val(Text3.Text) + Date
Label7.ForeColor = vbBlack
Command1.Enabled = True
End If
End Sub

