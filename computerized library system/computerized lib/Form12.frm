VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form12 
   BackColor       =   &H00808000&
   Caption         =   "Pay Fines"
   ClientHeight    =   4350
   ClientLeft      =   7050
   ClientTop       =   3420
   ClientWidth     =   4575
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   4350
   ScaleWidth      =   4575
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
      Left            =   360
      TabIndex        =   4
      Top             =   3840
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
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
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Format          =   76939265
      CurrentDate     =   39706
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Paid :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1455
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
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "O.R. Number:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text2.Text = "" Then
    MsgBox "All fields required not to be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
If Label4.Caption = "frmdam" Then
With Form1
    .Adodc5.Recordset.Fields(11) = Format(DTPicker1.Value, "mm/dd/yy")
    .Adodc5.Recordset.Fields(12) = (Text2.Text)
    .Adodc5.Recordset.Fields(9) = (Text3.Text)
    .Adodc5.Recordset.Fields(10) = "Returned"
    .Adodc5.Recordset.Update
    .Adodc5.Refresh
    .Adodc8.Recordset.Delete
    .Adodc8.Refresh
    End With
ElseIf Label4.Caption = "frmlost" Then
    With Form1
    .Adodc5.Recordset.Fields(11) = Format(DTPicker1.Value, "mm/dd/yy")
    .Adodc5.Recordset.Fields(12) = (Text2.Text)
    .Adodc5.Recordset.Fields(9) = (Text3.Text)
    .Adodc5.Recordset.Fields(10) = "Returned"
    .Adodc5.Recordset.Update
    .Adodc6.Recordset.Delete
    .Adodc6.Refresh
    End With
    Call ReturnBookQty
 Else
 With Form1
    .Adodc5.Recordset.Fields(11) = Format(DTPicker1.Value, "mm/dd/yy")
    .Adodc5.Recordset.Fields(12) = (Text2.Text)
    .Adodc5.Recordset.Fields(9) = (Text3.Text)
    .Adodc5.Recordset.Fields(10) = "Returned"
    .Adodc5.Recordset.Update
    .Adodc5.Refresh
    End With
    End If

MsgBox "Returned Book has been added.", vbInformation, "Library System"
Unload Me

End Sub

Private Sub Command2_Click()
DTPicker1.Value = Date
Text2.Text = ""

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label7.Caption = Form1.Adodc2.Recordset.Fields(6)
End Sub

Private Sub Label4_Change()
If Label4.Caption = "frmdam" Then
Text3.Text = (Form1.Adodc8.Recordset.Fields(9))
DTPicker1.Value = Format(Date, "mm/dd/yy")
ElseIf Label4.Caption = "frmlost" Then
Text3.Text = Val(Form1.Adodc6.Recordset.Fields(9)) + Val(Label7.Caption)
DTPicker1.Value = Format(Date, "mm/dd/yy")
Else
Text3.Text = (Form1.Adodc5.Recordset.Fields(9))
DTPicker1.Value = Format(Date, "mm/dd/yy")
End If
End Sub

Private Sub Text2_Change()
a = Val(Text2.Text)

End Sub

Private Sub Text3_Change()
b = Val(Text3.Text)
If Not (IsNumeric(b)) Then
MsgBox "Only integer Values are to be inputed in the field O.R. Number"
Else
End If
End Sub
