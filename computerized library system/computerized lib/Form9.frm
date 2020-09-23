VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00808000&
   Caption         =   "Damage Book"
   ClientHeight    =   4335
   ClientLeft      =   6645
   ClientTop       =   3615
   ClientWidth     =   4560
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   4335
   ScaleWidth      =   4560
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1800
      TabIndex        =   14
      Text            =   "0"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Text            =   "0"
      Top             =   1680
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
      Left            =   3000
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label13 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fines:"
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
      TabIndex        =   17
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fine For Damages:"
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
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Borrowed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Returned :"
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
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Days after Due:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fines :"
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
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning:Make sure of what you are                  doing because you cannot edit             it later."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3480
      Picture         =   "Form9.frx":0442
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   4575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    Form1.Adodc3.Recordset.Fields(7) = Format(Date, "mm/dd/yy")
    Form1.Adodc3.Recordset.Fields(8) = (Text9.Text)
    Form1.Adodc3.Recordset.Fields(9) = (Text3.Text)
    Form1.Adodc3.Recordset.Fields(10) = "Damaged"
    Form1.Adodc3.Recordset.Update
    Form1.Adodc3.Refresh
    Form1.Adodc8.Refresh
Call ReturnBookQty
MsgBox "Damage Book has been added to Damage Books.", vbInformation, "Library System"
Unload Me

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = (Form1.Adodc3.Recordset.Fields(5))
Label7.Caption = (Form1.Adodc3.Recordset.Fields(6))
Label13.Caption = Format(Date, "mm/dd/yy")
Call finesCharge_
If Date > DateValue(Label7.Caption) Then
Text10.Text = Date - DateValue(Label7.Caption)
Else
Text10.Text = 0
'Text2.SetFocus
End If
End Sub

Private Sub Text10_Change()
Text3.Text = Val(Text2.Text) + Val(Text10.Text)
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
Command1.Enabled = False
Label5.ForeColor = vbRed
ElseIf Not (IsNumeric(Text2.Text)) Then
MsgBox "Please Input A numeric Value on Field Fine for Damages"
Label5.ForeColor = vbRed
Command1.Enabled = False
Else
Command1.Enabled = True
Label5.ForeColor = vbBlack
Text3.Text = Val(Text2.Text) + Val(Text10.Text)
End If
End Sub

Private Sub Text3_Change()
Text3.Text = Val(Text2.Text) + Val(Text10.Text)
End Sub

Private Sub Text9_Change()
On Error Resume Next
'If Text9.Text = "" Then
    'Text9.Text = "0"
'End If
Text10.Text = finesCharge * Text9.Text

End Sub

