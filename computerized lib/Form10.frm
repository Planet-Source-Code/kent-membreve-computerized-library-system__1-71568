VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00808000&
   Caption         =   "Category"
   ClientHeight    =   5460
   ClientLeft      =   5865
   ClientTop       =   3420
   ClientWidth     =   9135
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   ScaleHeight     =   5460
   ScaleWidth      =   9135
   Begin MSDataGridLib.DataGrid DataGrid6 
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Frame Frame6 
      Caption         =   "Add Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "&EDIT"
         Height          =   350
         Left            =   4320
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
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
         Height          =   350
         Left            =   3000
         TabIndex        =   5
         Top             =   1680
         Width           =   1305
      End
      Begin VB.CommandButton Command24 
         Caption         =   "&Delete"
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
         Left            =   5520
         TabIndex        =   4
         Top             =   1680
         Width           =   1425
      End
      Begin VB.CommandButton Command25 
         Caption         =   "&Refresh"
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
         Left            =   6960
         TabIndex        =   3
         Top             =   1680
         Width           =   1425
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   10
         ScaleHeight     =   345
         ScaleWidth      =   2985
         TabIndex        =   2
         Top             =   1680
         Width           =   2980
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin VB.TextBox Text19 
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
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Category :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Version 1.0"
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
      Left            =   0
      TabIndex        =   10
      Top             =   3600
      Width           =   7935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   9135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "WELCOME TO THE COMPUTERIZED LIBRARY SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   4560
      Width           =   9135
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc6_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc6.Recordset.RecordCount) <= 0 Then
    Adodc6.Caption = "0"
    Text19.Text = ""
Else
    'Adodc6.Caption = (Adodc6.Recordset.AbsolutePosition)
    Text19.Text = (Adodc6.Recordset.Fields(0))
End If
End Sub

Private Sub Command1_Click()
Form13.Text2.Text = "EDIT"
Form13.Show vbModal
End Sub

Private Sub Command23_Click()
Form13.Text2.Text = "ADD"
Form13.Show vbModal
End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim repp3 As Integer
If Val(Adodc6.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp3 = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp3 = vbYes Then
    Adodc6.Recordset.Delete
    Adodc6.Recordset.MoveNext
    If Adodc6.Recordset.EOF Then
        Adodc6.Recordset.MoveLast
    End If
    Call ReturnBookQty
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If

End Sub

Private Sub Command25_Click()
Adodc6.Refresh
End Sub

Private Sub Form_Load()
Adodc6.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc6.RecordSource = "Select * From CATEGORY Order by CATEGORY"
        Set DataGrid6.DataSource = Adodc6
End Sub

Private Sub Form_Resize()
Form10.Width = 8145
Form10.Height = 6000
End Sub

