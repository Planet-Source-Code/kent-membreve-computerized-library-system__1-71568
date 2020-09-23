VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD BARROWER"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form2.frx":058A
      Left            =   1800
      List            =   "Form2.frx":059A
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "1ST YR."
      Top             =   2040
      Width           =   2055
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
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
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
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
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
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:If you want to Edit the existing             Barrower    just type the                      Barrower's ID bellow."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3360
      Picture         =   "Form2.frx":05C1
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "Form2.frx":0E8B
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Barrower's ID :"
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
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Barrower's Name :"
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
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Course :"
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
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Year :"
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
End
Attribute VB_Name = "Form2"
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "All fields required not to be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
If Command1.Caption = "&OK" Then
With Form1
    .Adodc1.Recordset.AddNew
    .Adodc1.Recordset.Fields(0) = (Text1.Text)
    .Adodc1.Recordset.Fields(1) = (Text2.Text)
    .Adodc1.Recordset.Fields(2) = (Text3.Text)
    .Adodc1.Recordset.Fields(3) = (Combo1.Text)
    .Adodc1.Recordset.Fields(4) = (Text4.Text)
    .Adodc1.Recordset.Fields(5) = (Text5.Text)
    .Adodc1.Recordset.Fields(6) = "Safe"
    .Adodc1.Recordset.Update
    .Adodc1.Refresh
    
End With
MsgBox "New Barrower has been added.", vbInformation, "Library System"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Text2.Text = ""
Text3.Text = ""
Combo1.Text = "1ST YR."
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim reply As Integer
reply = MsgBox("Are you Sure You want to Exit?", vbExclamation + vbYesNo, "Library System")
If reply = vbYes Then
Unload Me
Else
    Cancel = 1
End If
End Sub

Private Sub Text1_Change()
Text2.Locked = False
Text3.Locked = False
Combo1.Locked = False
Form1.Adodc1.Recordset.MoveFirst
Form1.Adodc1.Recordset.Find "BORROWERS_ID like '" & Text1.Text & "'"
If Form1.Adodc1.Recordset.EOF Then
    Command2_Click
    Command1.Caption = "&OK"
    Me.Caption = "ADD BARROWER"
    Label2.ForeColor = vbBlack
    Command1.Enabled = True
    Exit Sub
Else
    MsgBox "That Borrower's ID is Already Taken"
    Label2.ForeColor = vbRed
    Command1.Enabled = False
End If
End Sub



Private Sub Text2_Change()
If Text2.Text = "" Then
Label3.ForeColor = vbRed
Command1.Enabled = False
Else
Label3.ForeColor = vbBlack
Command1.Enabled = True
End If
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Then
Label4.ForeColor = vbRed
Command1.Enabled = False
Else
Label4.ForeColor = vbBlack
Command1.Enabled = True
End If

End Sub

Private Sub Text4_Change()
If Text4.Text = "" Then
Label1.ForeColor = vbRed
Command1.Enabled = False
Else
Label1.ForeColor = vbBlack
Command1.Enabled = True
End If

End Sub

Private Sub Text5_Change()
If Text5.Text = "" Then
Label6.ForeColor = vbRed
Command1.Enabled = False
Else
Label6.ForeColor = vbBlack
Command1.Enabled = True
End If

End Sub
