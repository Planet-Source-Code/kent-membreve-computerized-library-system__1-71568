VERSION 5.00
Begin VB.Form Dialog1 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Registration"
   ClientHeight    =   3195
   ClientLeft      =   6375
   ClientTop       =   3750
   ClientWidth     =   6015
   Icon            =   "Dialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pet's Name:"
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
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RE-TYPE PASSWORD:"
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
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   6015
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub OKButton_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox "All fields required not to be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
If frmLogin.Adodc1.Recordset.EOF Then
  If Text2.Text = Text3.Text Then
    With frmLogin
    
    .Adodc1.Recordset.AddNew
    .Adodc1.Recordset.Fields(0) = (Text2.Text)
    .Adodc1.Recordset.Fields(1) = (Text1.Text)
    .Adodc1.Recordset.Fields(2) = (Text4.Text)
    .Adodc1.Recordset.Fields(3) = "Attendant"
    .Adodc1.Recordset.Update
    End With
    MsgBox "New User has been added.", vbInformation, "Library System"
Unload Me
frmLogin.Show
Else
MsgBox "Your Password Must Match With the Re-type Password!"
Label3.ForeColor = vbRed
Label4.ForeColor = vbRed
End If
Else
MsgBox "Registration failed User Name Already exist"
End If

End Sub

Private Sub Text1_Change()
On Error Resume Next
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
frmLogin.Adodc1.Recordset.MoveFirst
frmLogin.Adodc1.Recordset.Find "USER_NAME like '" & Text1.Text & "'"
If frmLogin.Adodc1.Recordset.EOF Then
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
OKButton.Enabled = True
    Me.Caption = "ADD USER"
    Label1.ForeColor = vbBlack
    Exit Sub
Else
    MsgBox "That USER NAME Already Exist"
    Label1.ForeColor = vbRed
    OKButton.Enabled = False
End If
End Sub


Private Sub Text2_Change()
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub Text3_Change()
Label4.ForeColor = vbBlack
Label3.ForeColor = vbBlack
End Sub
