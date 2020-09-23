VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Dialog7 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Settings"
   ClientHeight    =   4665
   ClientLeft      =   7770
   ClientTop       =   4140
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8280
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   7920
      TabIndex        =   12
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Info"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Change Password"
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "New Password:"
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
         Left            =   2400
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Back"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Level:"
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
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      TabIndex        =   2
      Top             =   720
      Width           =   975
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Text2.Enabled = True
Text2.Text = ""
Command1.Visible = False
Label6.Visible = True
Label2.Caption = "Old Password:"
End Sub

Private Sub Command2_Click()
frmLogin.Adodc1.RecordSource = frmLogin.Adodc1.Recordset.EditMode
Frame1.Width = 4575
Frame1.Enabled = True
Command3.Enabled = True
Command2.Enabled = False
Text4.Text = ""
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Donot input null values to the fields"
Else
With frmLogin.Adodc1
.Recordset.Fields(0) = Text4.Text
.Recordset.Fields(1) = Text1.Text
.Recordset.Fields(2) = Text3.Text
.Recordset.Fields(3) = Label5.Caption
.Recordset.Update
Text2.Enabled = False
Text2.Text = Text4.Text
Label2.Caption = "Password"
Frame1.Width = 2415
Label6.Visible = False
Frame1.Enabled = False
Command1.Visible = True
Command3.Enabled = False
Command2.Enabled = True
End With
End If
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc1.RecordSource = "Select * From SECURITY_PASSWORD Order by USER_NAME"
        Set DataGrid1.DataSource = Adodc1
Adodc1.Recordset.Find "USER_NAME='" & (Form1.Label16.Caption) & "'"
Text1.Text = (Form1.Label16.Caption)
Text2.Text = (frmLogin.Adodc1.Recordset.Fields(0))
Text3.Text = (Adodc1.Recordset.Fields(2))
Label5.Caption = (Adodc1.Recordset.Fields(3))
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
Label1.ForeColor = vbRed
Command3.Enabled = False
ElseIf Text1.Text = Form1.Label16.Caption Then
Command3.Enabled = False
Else
Text2.Text = ""
Label1.ForeColor = vbBlack
Command3.Enabled = False
End If
End Sub


Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text2_LostFocus()
If Text2.Text <> Adodc1.Recordset.Fields(0) Then
MsgBox "Invalid old password"
Command3.Enabled = False
Label2.ForeColor = vbRed
Else
Command3.Enabled = True
Label2.ForeColor = vbBlack
End If
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Then
Label3.ForeColor = vbRed
Command3.Enabled = False
ElseIf Text3.Text = Adodc1.Recordset.Fields(2) Then
Command3.Enabled = False
Else
Label3.ForeColor = vbBlack
Command3.Enabled = True
End If
End Sub
