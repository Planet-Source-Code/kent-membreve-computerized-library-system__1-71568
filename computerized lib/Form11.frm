VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form11 
   BackColor       =   &H00808000&
   Caption         =   "Print"
   ClientHeight    =   5025
   ClientLeft      =   8430
   ClientTop       =   3825
   ClientWidth     =   3480
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   ScaleHeight     =   5025
   ScaleWidth      =   3480
   Begin TabDlg.SSTab SSTab1 
      Height          =   4860
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8573
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Borrowers"
      TabPicture(0)   =   "Form11.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Cmd23"
      Tab(0).Control(2)=   "DataCombo1"
      Tab(0).Control(3)=   "Image1"
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(5)=   "Label2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Books"
      TabPicture(1)   =   "Form11.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "DataCombo2"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Image2"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Borrowed Books"
      TabPicture(2)   =   "Form11.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command2"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "Image3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Print By Book Status"
      TabPicture(3)   =   "Form11.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Image4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "DataCombo3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Command3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame3"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame3 
         Caption         =   "Print Option"
         Height          =   855
         Left            =   720
         TabIndex        =   19
         Top             =   1080
         Width           =   2655
         Begin VB.OptionButton Option6 
            Caption         =   "Print All Book Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Print By Status"
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
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print Preview"
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   3000
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Select"
      End
      Begin VB.Frame Frame1 
         Caption         =   "Print Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74220
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Print by  Year"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Print all Borrowers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.CommandButton Cmd23 
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74220
         TabIndex        =   6
         Top             =   3240
         Width           =   2625
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74220
         TabIndex        =   5
         Top             =   3240
         Width           =   2625
      End
      Begin VB.Frame Frame2 
         Caption         =   "Print Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74220
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
         Begin VB.OptionButton Option3 
            Caption         =   "Print all Books"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Print by Author"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74100
         TabIndex        =   1
         Top             =   2040
         Width           =   2625
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   -74220
         TabIndex        =   10
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   -74220
         TabIndex        =   11
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Book Status:"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image Image4 
         Height          =   720
         Left            =   2520
         Picture         =   "Form11.frx":04B2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   840
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   -72420
         Picture         =   "Form11.frx":0A3C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "PRINT BARROWERS"
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
         Left            =   -74220
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Year:"
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
         Left            =   -74220
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PRINT BOOKS"
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
         Left            =   -74220
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   -72420
         Picture         =   "Form11.frx":0FC6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Author:"
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
         Left            =   -74220
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "PRINT BORROWED BOOKS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74100
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   -72300
         Picture         =   "Form11.frx":1550
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   840
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd23_Click()
If Option2.Value = True Then
    DE1.rsCommand1.Filter = ""
End If
If Option1.Value = True Then
    If DataCombo1.Text = "" Then
        MsgBox "Pls. select a valid year for barrowers.", vbExclamation, "Library System"
    Exit Sub
    End If
    DE1.rsCommand1.Filter = "CURRENT_YEAR ='" & (DataCombo1.Text) & "'"
End If
'DE1.rsCommand1.Close
Rpt1.Show

End Sub

Private Sub Command1_Click()
DE1.rsCommand2.Filter = ""
If Option3.Value = True Then
    DE1.rsCommand2.Filter = ""
End If
If Option4.Value = True Then
    If DataCombo2.Text = "" Then
        MsgBox "Pls. select a valid Author.", vbExclamation, "Library System"
    Exit Sub
    End If
    DE1.rsCommand2.Filter = "AUTHOR ='" & (DataCombo2.Text) & "'"
End If
'DE1.rsCommand2.Close
Rpt2.Show
End Sub

Private Sub Command2_Click()
'DE1.rsCommand3.Close
Rpt3.Show
End Sub

Private Sub Command23_Click()
End Sub

Private Sub Command3_Click()
DE1.rsCommand3.Filter = ""
If Option6.Value = True Then
    DE1.rsCommand3.Filter = ""
End If
If Option5.Value = True Then
    If DataCombo3.Text = "Select" Then
        MsgBox "Pls. select a valid Book Status.", vbExclamation, "Library System"
    Exit Sub
    End If
    DE1.rsCommand3.Filter = "BOOK_STATUS ='" & (DataCombo3.Text) & "'"
End If
DataReport1.Show
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
SSTab1.Tab = 0
Option1.Value = True
Option4.Value = True
Option5.Value = True
Call yearPrint
authorPrint
Stat
End Sub

Private Sub Form_Resize()
Form11.Height = 5565
Form11.Width = 3720
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Label2.Visible = True
    DataCombo1.Visible = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    Label2.Visible = False
    DataCombo1.Visible = False
End If
End Sub

Private Sub Option3_Click()
If Option4.Value = False Then
    Label4.Visible = False
    DataCombo2.Visible = False
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    Label4.Visible = True
    DataCombo2.Visible = True
End If
End Sub


Private Sub Option5_Click()
If Option5.Value = True Then
    Label6.Visible = True
    DataCombo3.Visible = True
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
    Label6.Visible = False
    DataCombo3.Visible = False
End If

End Sub
