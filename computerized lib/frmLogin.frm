VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3555
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4065
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100.411
   ScaleMode       =   0  'User
   ScaleWidth      =   3816.815
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5520
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   735
      Left            =   5520
      TabIndex        =   17
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   2160
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   1455
      Left            =   4680
      TabIndex        =   14
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2566
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
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Just input the Name of your pet on this field"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "forgot pass"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   -120
      TabIndex        =   11
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pet's Name"
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
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Version 1.5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   -120
      TabIndex        =   10
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   -120
      TabIndex        =   12
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
Text3.Text = (Adodc1.Recordset.Fields(3))
End Sub

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
If Adodc1.Recordset.EOF Then
MsgBox "That User is Not Yet Registered"
Else
    If txtPassword = (Adodc1.Recordset.Fields(0)) Then
        LoginSucceeded = True
        Me.Hide
        frmSplash.Show
    Else
        MsgBox "Invalid Password or User Name, try again!", , "Login"
        txtPassword.SetFocus
        Label1.Visible = True
        Text1.Visible = True
        Command1.Visible = True
    End If
    End If
End Sub

Private Sub Command1_Click()
 If Text1.Text = (Adodc1.Recordset.Fields(2)) Then
 Dialog.Show
 Unload Me
 Else
 
MsgBox "Wrong Pet's Name"
End If
End Sub

Private Sub Command2_Click()
Dialog2.Show
Unload Me
End Sub

Private Sub DataGrid1_Change()
Text2.Text = (Adodc1.Recordset.Fields(1))
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc1.RecordSource = "Select * From SECURITY_PASSWORD Order by USER_NAME"
        Set DataGrid1.DataSource = Adodc1
    Adodc2.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc2.RecordSource = "Select * From themes Order by themes"
        Set DataGrid2.DataSource = Adodc2
        Label2.Caption = (Adodc2.Recordset.Fields(0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim reply As Integer
reply = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo, "Library System")
If reply = vbYes Then
    End
Else
    Cancel = 1
End If
End Sub

Private Sub Label2_Change()
If Label2.Caption = "Pandan" Then
Form1.Label18.BackColor = &HC0FFC0
Form1.Label20.BackColor = &HC0FFC0
Form1.Label4.BackColor = &HC0FFC0
Form1.Label8.BackColor = &HC0FFC0
Form1.Label7.BackColor = &HC0FFC0
Form1.Label9.BackColor = &HC0FFC0
Form1.Label10.BackColor = &HC0FFC0
Form1.Label12.BackColor = &HC0FFC0
Form1.Label71.BackColor = &HC0FFC0
Form1.Option1.BackColor = &HC0FFC0
Form1.Option2.BackColor = &HC0FFC0
Form1.pnp.Checked = False
Form1.redapple.Checked = False
Form1.dalandan.Checked = False
Form1.pandan.Checked = True
Form1.banana.Checked = False
Form1.choco.Checked = False
Form1.Emo.Checked = False
ElseIf Label2.Caption = "Red" Then
Form1.Label18.BackColor = &HFF&
Form1.Label20.BackColor = &HFF&
Form1.Label4.BackColor = &HFF&
Form1.Label8.BackColor = &HFF&
Form1.Label7.BackColor = &HFF&
Form1.Label9.BackColor = &HFF&
Form1.Label10.BackColor = &HFF&
Form1.Label12.BackColor = &HFF&
Form1.Label71.BackColor = &HFF&
Form1.Option1.BackColor = &HFF&
Form1.Option2.BackColor = &HFF&
Form1.pnp.Checked = False
Form1.redapple.Checked = True
Form1.dalandan.Checked = False
Form1.pandan.Checked = False
Form1.banana.Checked = False
Form1.choco.Checked = False
Form1.Emo.Checked = False
ElseIf Label2.Caption = "Emo" Then
Form1.Label18.BackColor = &H0&
Form1.Label20.BackColor = &H0&
Form1.Label4.BackColor = &H0&
Form1.Label8.BackColor = &H0&
Form1.Label7.BackColor = &H0&
Form1.Label9.BackColor = &H0&
Form1.Label10.BackColor = &H0&
Form1.Label12.BackColor = &H0&
Form1.Label71.BackColor = &H0&
Form1.Option1.BackColor = &H0&
Form1.Option2.BackColor = &H0&
Form1.pnp.Checked = False
Form1.redapple.Checked = False
Form1.dalandan.Checked = False
Form1.pandan.Checked = False
Form1.banana.Checked = False
Form1.choco.Checked = False
Form1.Emo.Checked = True
ElseIf Label2.Caption = "pnp" Then
Form1.Label18.BackColor = &H80FF&
Form1.Label20.BackColor = &H80FF&
Form1.Label4.BackColor = &H80FF&
Form1.Label8.BackColor = &H80FF&
Form1.Label7.BackColor = &H80FF&
Form1.Label9.BackColor = &H80FF&
Form1.Label10.BackColor = &H80FF&
Form1.Label12.BackColor = &H80FF&
Form1.Label71.BackColor = &H80FF&
Form1.Option1.BackColor = &H80FF&
Form1.Option2.BackColor = &H80FF&
Form1.pnp.Checked = True
Form1.redapple.Checked = False
Form1.dalandan.Checked = False
Form1.pandan.Checked = False
Form1.banana.Checked = False
Form1.choco.Checked = False
Form1.Emo.Checked = False
ElseIf Label2.Caption = "dalandan" Then
Form1.Label18.BackColor = &HFF00&
Form1.Label20.BackColor = &HFF00&
Form1.Label4.BackColor = &HFF00&
Form1.Label8.BackColor = &HFF00&
Form1.Label7.BackColor = &HFF00&
Form1.Label9.BackColor = &HFF00&
Form1.Label10.BackColor = &HFF00&
Form1.Label12.BackColor = &HFF00&
Form1.Label71.BackColor = &HFF00&
Form1.Option1.BackColor = &HFF00&
Form1.Option2.BackColor = &HFF00&
Form1.pnp.Checked = False
Form1.redapple.Checked = False
Form1.dalandan.Checked = True
Form1.pandan.Checked = False
Form1.banana.Checked = False
Form1.choco.Checked = False
Form1.Emo.Checked = False
ElseIf Label2.Caption = "banana" Then
Form1.Label18.BackColor = &HFFFF&
Form1.Label20.BackColor = &HFFFF&
Form1.Label4.BackColor = &HFFFF&
Form1.Label8.BackColor = &HFFFF&
Form1.Label7.BackColor = &HFFFF&
Form1.Label9.BackColor = &HFFFF&
Form1.Label10.BackColor = &HFFFF&
Form1.Label12.BackColor = &HFFFF&
Form1.Label71.BackColor = &HFFFF&
Form1.Option1.BackColor = &HFFFF&
Form1.Option2.BackColor = &HFFFF&
Form1.pnp.Checked = False
Form1.redapple.Checked = False
Form1.dalandan.Checked = False
Form1.pandan.Checked = False
Form1.banana.Checked = True
Form1.choco.Checked = False
Form1.Emo.Checked = False
Else
Form1.Label18.BackColor = &H8F&
Form1.Label20.BackColor = &H8F&
Form1.Label4.BackColor = &H8F&
Form1.Label8.BackColor = &H8F&
Form1.Label7.BackColor = &H8F&
Form1.Label9.BackColor = &H8F&
Form1.Label10.BackColor = &H8F&
Form1.Label12.BackColor = &H8F&
Form1.Label71.BackColor = &H8F&
Form1.Option1.BackColor = &H8F&
Form1.Option2.BackColor = &H8F&
Form1.pnp.Checked = False
Form1.redapple.Checked = False
Form1.dalandan.Checked = False
Form1.pandan.Checked = False
Form1.banana.Checked = False
Form1.choco.Checked = True
Form1.Emo.Checked = False
End If
End Sub
Private Sub txtUserName_Change()
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "USER_NAME like '" & (txtUserName.Text) & "'"
Text3.Refresh
End Sub
