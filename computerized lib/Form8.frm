VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00808000&
   Caption         =   "Found Book"
   ClientHeight    =   3885
   ClientLeft      =   7050
   ClientTop       =   3825
   ClientWidth     =   4530
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   3885
   ScaleWidth      =   4530
   Begin VB.TextBox Text9 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text10 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
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
      TabIndex        =   1
      Top             =   3240
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
      Left            =   2880
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1920
      TabIndex        =   14
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3480
      Picture         =   "Form8.frx":0442
      Top             =   0
      Width           =   480
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
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Days After reported Lost :"
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
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
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
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Reported Lost"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    Form1.Adodc6.Recordset.Fields(7) = Format(Date, "mm/dd/yy")
    Form1.Adodc6.Recordset.Fields(8) = (Text9.Text)
    Form1.Adodc6.Recordset.Fields(9) = (Text10.Text)
    Form1.Adodc6.Recordset.Fields(10) = "Returned"
    Form1.Adodc6.Recordset.Update
    Form1.Adodc6.Refresh
    Form1.Adodc5.Refresh
    Form1.Adodc2.Recordset.Fields(8) = Form1.Adodc2.Recordset.Fields(8) - 1
    Form1.Adodc2.Recordset.Fields(9) = Form1.Adodc2.Recordset.Fields(9) + 1
    Form1.Adodc2.Recordset.Update
    Form1.Adodc2.Refresh
MsgBox "Returned Book has been added.", vbInformation, "Library System"
Unload Me
Call ReturnBookQty
MsgBox "Lost Book has been added to Lost Books Records.", vbInformation, "Library System"
Me.Hide

End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Label5.Caption = (Form1.Adodc6.Recordset.Fields(7))
Call finesCharge_
Text1.Text = Format(Date, "mm/dd/yy")
If Date > DateValue(Label5.Caption) Then
Text9.Text = Date - DateValue(Label5.Caption)
Else
Text9.Text = 0
End If
End Sub

Private Sub Form_Resize()
Form8.Width = 4770
Form8.Height = 4425
End Sub

Private Sub Text9_Change()
On Error Resume Next
If Text9.Text = "" Then
    Text9.Text = "0"
End If
Text10.Text = finesCharge * Val(Text9.Text)

End Sub


