VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Category"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2835
   Icon            =   "à.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1215
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command2.Caption = "&Add" Then
If Text1.Text = "" Then
    MsgBox "Category cannot be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
With Form10
    .Adodc6.Recordset.AddNew
    .Adodc6.Recordset.Fields(0) = (Text1.Text)
    .Adodc6.Recordset.Update
End With
MsgBox "New Category has been successfully added.", vbInformation, "Library System"
Unload Me
Else
If Text1.Text = "" Then
    MsgBox "Category cannot be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
With Form10
.Adodc6.RecordSource = Form10.Adodc6.Recordset.EditMode
.Adodc6.Recordset.Fields(0) = (Text1.Text)
.Adodc6.Recordset.Update
End With
MsgBox "Category has been successfully Updated.", vbInformation, "Library System"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text2_Change()
If Text2.Text = "ADD" Then
Form13.Caption = "ADD"
Command1.Caption = "ADD"
Else
Command1.Caption = "UPDATE"
Form13.Caption = "EDIT"
Form10.Adodc6.RecordSource = Form10.Adodc6.Recordset.EditMode
Text1.Text = (Form10.Adodc6.Recordset.Fields(0))
End If


End Sub
