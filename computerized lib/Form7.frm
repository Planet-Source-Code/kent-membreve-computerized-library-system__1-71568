VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00808000&
   Caption         =   "Lost Books"
   ClientHeight    =   3840
   ClientLeft      =   7245
   ClientTop       =   3825
   ClientWidth     =   4545
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   3840
   ScaleWidth      =   4545
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1200
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
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
      Left            =   2760
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
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
      Left            =   1440
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   840
      Width           =   2055
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
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Date lost :"
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
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Days After Due :"
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2280
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
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3480
      Picture         =   "Form7.frx":030A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    Form1.Adodc3.Recordset.Fields(7) = Format(Date, "mm/dd/yy")
    Form1.Adodc3.Recordset.Fields(8) = (Text9.Text)
    Form1.Adodc3.Recordset.Fields(9) = (Text10.Text)
    Form1.Adodc3.Recordset.Fields(10) = "Lost"
    Form1.Adodc3.Recordset.Update
    Form1.Adodc3.Refresh
    Form1.Adodc5.Refresh
    MsgBox "Lost Book has been added to Lost Books Records.", vbInformation, "Library System"
Unload Me
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Me.Hide
frmLogin.Adodc1.Refresh
End Sub

Private Sub Form_Load()
Label5.Caption = (Form1.Adodc3.Recordset.Fields(5))
Call finesCharge_
Label6.Caption = (Form1.Adodc3.Recordset.Fields(6))
Text1.Text = Format(Date, "mm/dd/yy")
If Date > DateValue(Label6.Caption) Then
Text9.Text = Date - DateValue(Label6.Caption)
Else
Text9.Text = 0
End If
End Sub

Private Sub Form_Resize()
Form7.Height = 4380
Form7.Width = 4785
End Sub

Private Sub Text9_Change()
On Error Resume Next
If Text9.Text = "" Then
    Text9.Text = "0"
End If
Text10.Text = finesCharge * Val(Text9.Text)

End Sub

