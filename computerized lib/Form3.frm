VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD BOOK"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   855
      Left            =   5040
      TabIndex        =   20
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5400
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text7 
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
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   4
      Top             =   2400
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
      TabIndex        =   10
      Top             =   4200
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
      TabIndex        =   9
      Top             =   4200
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
      TabIndex        =   8
      Top             =   4200
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
      Top             =   2040
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
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Borrowed:"
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
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
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
      TabIndex        =   22
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "ADD"
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity :"
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
      TabIndex        =   18
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
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
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Published :"
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
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category :"
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
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:Red Labels Means Wrong inputs! Correct these Errors or you will not be able to continue with the Adding of books."
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
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3360
      Picture         =   "Form3.frx":058A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   0
      Picture         =   "Form3.frx":0E54
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book NO :"
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
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN :"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title :"
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
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Author :"
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
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y As Integer
Private Sub Combo1_Change()
If Combo1.Text = "" Then
    Combo1.Text = "1ST YR."
End If
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Label10 = "ADD" Then
Text1.Text = "ICA" & (Adodc1.Recordset.Fields(0))
Else
Exit Sub
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text = "" Then
Command1.Enabled = False
ElseIf Not (IsNumeric(Combo3.Text)) Then
MsgBox "must input a numeric value"
Label6.ForeColor = vbRed
Command1.Enabled = False
ElseIf Len(Combo3.Text) > 4 Then
MsgBox "must input only up to 4 integers"
Label6.ForeColor = vbRed
Command1.Enabled = False
Else
Label6.ForeColor = vbBlack
Command1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo3.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or DataCombo1.Text = "" Then
    MsgBox "All fields required not to be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
If Command1.Caption = "&OK" Then
With Form1
    .Adodc2.Recordset.AddNew
    .Adodc2.Recordset.Fields(0) = (Text1.Text)
    .Adodc2.Recordset.Fields(1) = (Text2.Text)
    .Adodc2.Recordset.Fields(2) = (Text3.Text)
    .Adodc2.Recordset.Fields(3) = (Text4.Text)
    .Adodc2.Recordset.Fields(4) = (DataCombo1.Text)
    .Adodc2.Recordset.Fields(5) = (Combo3.Text)
    .Adodc2.Recordset.Fields(6) = (Text6.Text)
    .Adodc2.Recordset.Fields(7) = (Text7.Text)
    .Adodc2.Recordset.Fields(8) = 0
    .Adodc2.Recordset.Fields(9) = (Text7.Text)
    .Adodc2.Recordset.Fields(10) = "Available"
    .Adodc2.Recordset.Update
    Adodc1.Recordset.Fields(0) = (Adodc1.Recordset.Fields(0)) + 1
    Adodc1.Recordset.Update
End With
MsgBox "New Book has been added.", vbInformation, "Library System"
Unload Me
End If
If Command1.Caption = "&Update" Then
    With Form1
    .Adodc2.Recordset.Fields(0) = (Text1.Text)
    .Adodc2.Recordset.Fields(1) = (Text2.Text)
    .Adodc2.Recordset.Fields(2) = (Text3.Text)
    .Adodc2.Recordset.Fields(3) = (Text4.Text)
    .Adodc2.Recordset.Fields(4) = (DataCombo1.Text)
    .Adodc2.Recordset.Fields(5) = (Combo3.Text)
    .Adodc2.Recordset.Fields(6) = (Text6.Text)
    .Adodc2.Recordset.Fields(7) = (Text7.Text)
    .Adodc2.Recordset.Fields(8) = (Form1.Adodc2.Recordset.Fields(8))
    .Adodc2.Recordset.Fields(9) = (Text7.Text)
    .Adodc2.Recordset.Fields(10) = (Form1.Adodc2.Recordset.Fields(10))
    .Adodc2.Recordset.Update
End With
MsgBox "Changes has been successfully save.", vbInformation, "Library System"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
DataCombo1.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Text2.SetFocus
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc1.RecordSource = "Select * From Autonum Order by AUOTO_NUM"
        Set DataGrid1.DataSource = Adodc1
        Adodc1.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
 Connect
    
    Set catRS = New ADODB.Recordset
        
    Y = 1601
    While Y <= 9999
        Combo3.AddItem Y
        Y = Y + 1
    Wend

Call BookCategory
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim reply As Integer
reply = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo, "Library System")
If reply = vbYes Then
    Unload Me
Else
    Cancel = 1
End If
End Sub

Private Sub Label10_Change()
If Label10.Caption = "EDIT" Then
Form3.Caption = "EDIT"
Text1.Text = Form1.Adodc2.Recordset.Fields(0)
Text2.Text = Form1.Adodc2.Recordset.Fields(1)
Text3.Text = Form1.Adodc2.Recordset.Fields(2)
Text4.Text = Form1.Adodc2.Recordset.Fields(3)
DataCombo1.Text = Form1.Adodc2.Recordset.Fields(4)
Combo3.Text = Form1.Adodc2.Recordset.Fields(5)
Text6.Text = Form1.Adodc2.Recordset.Fields(6)
Text7.Text = Form1.Adodc2.Recordset.Fields(7)
Text5.Text = Form1.Adodc2.Recordset.Fields(9)
Text8.Text = Form1.Adodc2.Recordset.Fields(8)
Text5.Visible = True
Text8.Visible = True
Label11.Visible = True
Label12.Visible = True
Command1.Top = 5040
Command1.Left = 120
Command2.Top = 5040
Command2.Left = 1440
Command3.Top = 5040
Command3.Left = 2760
Command1.Caption = "&Update"
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
DataCombo1.Locked = False
Combo3.Locked = False
Text6.Locked = False
Text7.Locked = False
Text5.Locked = False
Text8.Locked = False
End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Text1.Text = "" Then
Label2.ForeColor = vbRed
Command1.Enabled = False
Else
Label2.ForeColor = vbBlack
Command1.Enabled = True
End If
End Sub

Private Sub Text2_LostFocus()
On Error Resume Next

If Label10.Caption = "ADD" Then
Form1.Adodc2.Recordset.MoveFirst
Form1.Adodc2.Recordset.Find "ISBN like '" & Text2.Text & "'"
If Form1.Adodc2.Recordset.EOF Then
    Command1.Caption = "&OK"
    Me.Caption = "ADD BOOK"
    Command1.Enabled = True
    Label3.ForeColor = vbBlack
    Text1.Locked = False
    Text3.Locked = False
    Text4.Locked = False
    Text5.Locked = False
    Text6.Locked = False
    Text7.Locked = False
    DataCombo1.Locked = False
    Exit Sub
Else
    MsgBox "That ISBN Already Exist"
    Command1.Enabled = False
    Label3.ForeColor = vbRed
    Text1.Locked = True
    Text3.Locked = True
    Text4.Locked = True
    Text5.Locked = True
    Text6.Locked = True
    Text7.Locked = True
    DataCombo1.Locked = True
End If
Else
Exit Sub
End If
End Sub

Private Sub Text5_Change()
If Not (IsNumeric(Text5.Text)) Then
Label11.ForeColor = vbRed
Command1.Enabled = False
Else
Label11.ForeColor = vbBlack
Command1.Enabled = True
End If
End Sub

Private Sub Text6_Change()
On Error Resume Next
If Not (IsNumeric(Text6.Text)) Then
Label7.ForeColor = vbRed
Command1.Enabled = False
Else
Label7.ForeColor = vbBlack
Command1.Enabled = True
End If
End Sub

Private Sub Text7_Change()
On Error Resume Next
If Not (IsNumeric(Text7.Text)) Then
Label8.ForeColor = vbRed
Command1.Enabled = False
Else
Label8.ForeColor = vbBlack
Command1.Enabled = True
Text5.Text = Val(Text7.Text) - Val(Text8.Text)
End If
End Sub

Private Sub Text8_Change()
If Not (IsNumeric(Text8.Text)) Then
Label12.ForeColor = vbRed
Command1.Enabled = False
Else
Label12.ForeColor = vbBlack
Command1.Enabled = True
Text5.Text = Val(Text7.Text) - Val(Text8.Text)
End If
End Sub
